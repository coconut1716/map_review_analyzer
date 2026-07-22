from __future__ import annotations

import argparse
import atexit
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


TOOLS = Path(__file__).resolve().parent
ROOT = TOOLS
DEFAULT_OUT_ROOT = ROOT / "outputs" / "review_pipeline"
ANSI_RE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")


class ConsoleDashboard:
    """Render pipeline progress in-place, like a small console game screen."""

    STAGES = ("음식점 목록 수집", "카카오맵 리뷰 수집", "정량 순위표 생성")

    def __init__(self, region: str, run_dir: Path, *, plain: bool = False) -> None:
        self.region = region
        self.run_dir = run_dir
        self.enabled = not plain and sys.stdout.isatty() and self._enable_virtual_terminal()
        self.stage_states = ["pending"] * len(self.STAGES)
        self.restaurant_total = 0
        self.restaurant_done = 0
        self.current_index = 0
        self.current_item = "준비 중"
        self.detail = "실행 환경을 확인하고 있습니다."
        self.started_at = time.monotonic()
        self.completed = False
        self.result_path = ""
        self.result_files: list[str] = []
        self._entered = False

    @staticmethod
    def _enable_virtual_terminal() -> bool:
        if os.name != "nt":
            return True
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32
            handle = kernel32.GetStdHandle(-11)
            mode = ctypes.c_uint32()
            if handle in (0, -1) or not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                return False
            return bool(kernel32.SetConsoleMode(handle, mode.value | 0x0004))
        except Exception:
            return False

    @staticmethod
    def _clean(text: str) -> str:
        return ANSI_RE.sub("", text).replace("\r", " ").replace("\n", " ").strip()

    @staticmethod
    def _shorten(text: str, width: int) -> str:
        text = ConsoleDashboard._clean(text)
        if len(text) <= width:
            return text
        return text[: max(1, width - 1)] + "…"

    @staticmethod
    def _field_lines(label: str, value: str, width: int) -> list[str]:
        prefix = f"  {label:<7} "
        continuation = " " * len(prefix)
        available = max(20, width - len(prefix))
        clean = ConsoleDashboard._clean(value)
        chunks = [clean[index : index + available] for index in range(0, len(clean), available)] or [""]
        return [prefix + chunks[0], *(continuation + chunk for chunk in chunks[1:])]

    def start(self) -> None:
        if not self.enabled or self._entered:
            return
        self._entered = True
        sys.stdout.write("\x1b[2J\x1b[H\x1b[?25l")
        self.render()

    def begin_stage(self, stage: int, label: str) -> None:
        self.stage_states[stage - 1] = "running"
        self.current_item = label
        self.detail = "단계를 시작했습니다."
        self.render()

    def finish_stage(self, stage: int) -> None:
        self.stage_states[stage - 1] = "done"
        if stage == 2 and self.restaurant_total:
            self.restaurant_done = self.restaurant_total
        self.detail = "단계가 완료되었습니다."
        self.render()

    def complete(self, result_path: Path, result_files: list[Path]) -> None:
        self.completed = True
        self.result_path = str(result_path)
        self.result_files = [Path(file).name for file in result_files]
        self.current_item = "모든 단계 완료"
        self.detail = "XLSX, HTML, CSV를 생성했습니다."
        self.render()

    def fail_stage(self, stage: int, detail: str) -> None:
        self.stage_states[stage - 1] = "failed"
        self.detail = self._clean(detail)
        self.render()

    def consume(self, line: str) -> None:
        clean = self._clean(line)
        if not clean:
            return
        self.detail = clean

        match = re.search(r"grid\s+(\d+)\s*/\s*(\d+)", clean)
        if match:
            self.current_item = f"검색 격자 {match.group(1)} / {match.group(2)}"

        match = re.search(r"수집 식당 수:\s*(\d+)", clean)
        if match:
            self.restaurant_total = int(match.group(1))
            self.current_item = f"식당 {self.restaurant_total}개 발견"

        match = re.search(r"XLSX URL\s+(\d+)개", clean)
        if match:
            self.restaurant_total = int(match.group(1))
            self.current_item = "리뷰 수집 준비 완료"

        match = re.match(r"(\d+)\s*/\s*(\d+)\s+(https?://\S+)", clean)
        if match:
            self.current_index = int(match.group(1))
            self.restaurant_total = int(match.group(2))
            self.restaurant_done = max(self.restaurant_done, self.current_index - 1)
            place_id = re.search(r"place\.map\.kakao\.com/(\d+)", match.group(3))
            current = place_id.group(1) if place_id else match.group(3)
            self.current_item = f"식당 {self.current_index}: Kakao place {current}"

        if clean.startswith("reviews:") and self.current_index:
            self.restaurant_done = max(self.restaurant_done, self.current_index)

        self.render()

    def _progress_lines(self, width: int) -> list[str]:
        total = self.restaurant_total
        if total <= 0:
            return ["  진행도  □  식당 목록을 계산하는 중"]

        block_total = max(1, (total + 9) // 10)
        block_done = min(block_total, self.restaurant_done // 10)
        if self.restaurant_done >= total:
            block_done = block_total
        blocks = "■" * block_done + "□" * (block_total - block_done)
        chunk_size = max(10, min(50, width - 6))
        chunks = [blocks[index : index + chunk_size] for index in range(0, len(blocks), chunk_size)]
        lines = [f"  진행도  {self.restaurant_done:>4} / {total:<4}  (■ 1칸 = 식당 10개)"]
        lines.extend(f"          {chunk}" for chunk in chunks)
        return lines

    def render(self) -> None:
        if not self.enabled:
            return
        width = max(64, min(110, shutil.get_terminal_size((88, 24)).columns))
        elapsed = int(time.monotonic() - self.started_at)
        stage_icons = {"pending": "□", "running": "▶", "done": "■", "failed": "!"}
        lines = [
            "═" * width,
            "  KAKAO REVIEW LOADING",
            f"  지역    {self._shorten(self.region, width - 10)}",
            f"  경과    {elapsed // 60:02d}:{elapsed % 60:02d}",
            "─" * width,
        ]
        for index, (name, state) in enumerate(zip(self.STAGES, self.stage_states), start=1):
            lines.append(f"  {stage_icons[state]} {index}/3  {name}")

        if self.completed:
            lines.extend(["─" * width, "  ■ 완료", "  XLSX, HTML, CSV를 생성했습니다."])
            lines.extend(self._field_lines("저장 위치", self.result_path, width))
            lines.extend(self._field_lines("파일 이름", " / ".join(self.result_files), width))
            lines.extend(["─" * width, "  위 저장 폴더에서 결과 파일을 확인할 수 있습니다.", "═" * width])
        else:
            lines.extend(["─" * width, *self._progress_lines(width), "─" * width])
            lines.append(f"  현재    {self._shorten(self.current_item, width - 10)}")
            lines.append(f"  상태    {self._shorten(self.detail, width - 10)}")
            lines.extend(["─" * width, "  중단: Ctrl+C   상세 기록: 실행 폴더의 pipeline.log", "═" * width])
        sys.stdout.write("\x1b[H" + "\n".join(lines) + "\x1b[J")
        sys.stdout.flush()

    def close(self) -> None:
        if not self.enabled or not self._entered:
            return
        self.render()
        sys.stdout.write("\x1b[?25h\n")
        sys.stdout.flush()
        self._entered = False

def safe_name(value: str) -> str:
    text = "".join(ch if ch.isalnum() else "_" for ch in value.strip())
    return "_".join(part for part in text.split("_") if part) or "region"



def run_step(
    label: str,
    command: list[str],
    *,
    stage: int,
    dashboard: ConsoleDashboard,
    log_path: Path,
    env: dict[str, str] | None = None,
) -> None:
    display_command = " ".join(f'"{part}"' if " " in part else part for part in command)
    dashboard.begin_stage(stage, label)
    if not dashboard.enabled:
        print(f"\n[{label}]")
        print(display_command)

    child_env = (env or os.environ.copy()).copy()
    child_env.setdefault("PYTHONUNBUFFERED", "1")
    child_env.setdefault("PYTHONUTF8", "1")
    with log_path.open("a", encoding="utf-8") as log:
        log.write(f"\n[{datetime.now().isoformat(timespec='seconds')}] {label}\n{display_command}\n")
        process = subprocess.Popen(
            command,
            cwd=ROOT,
            env=child_env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        last_line = ""
        try:
            assert process.stdout is not None
            for line in process.stdout:
                last_line = line.rstrip("\r\n")
                log.write(line)
                log.flush()
                if dashboard.enabled:
                    dashboard.consume(last_line)
                else:
                    print(last_line)
            return_code = process.wait()
        except KeyboardInterrupt:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
            dashboard.fail_stage(stage, "사용자가 실행을 중단했습니다.")
            raise

    if return_code != 0:
        detail = last_line or f"exit code {return_code}"
        dashboard.fail_stage(stage, detail)
        raise SystemExit(f"{label} 실패: exit code {return_code}. 상세 기록: {log_path}")
    dashboard.finish_stage(stage)


def latest_xlsx(folder: Path, pattern: str) -> Path:
    files = sorted(folder.glob(pattern), key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        raise SystemExit(f"생성된 XLSX를 찾지 못했습니다: {folder / pattern}")
    return files[0]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="지역명 하나로 카카오맵 음식점 수집, 리뷰 수집, 정량 리뷰 순위표 생성을 이어서 실행합니다."
    )
    parser.add_argument("region", nargs="?", help="검색할 지역. 비우면 실행 중에 물어봅니다.")
    parser.add_argument("--api-key", default=os.getenv("KAKAO_REST_API_KEY"), help="Kakao REST API 키")
    parser.add_argument("--out-root", type=Path, default=DEFAULT_OUT_ROOT, help="실행 결과를 모을 상위 폴더")
    parser.add_argument("--mode", choices=["keyword", "category", "both", "grid"], default="grid")
    parser.add_argument("--radius", type=int, default=2000)
    parser.add_argument("--outer-radius", type=int, default=300, help="grid 모드 최종 기준점 반경(m)")
    parser.add_argument("--grid-step", type=int, default=100, help="grid 모드 검색 중심 좌표 간격(m)")
    parser.add_argument("--cell-radius", type=int, default=120, help="grid 모드 각 격자점 검색 반경(m)")
    parser.add_argument("--minimal", action="store_true", help="식당 목록 CSV/XLSX에 핵심 필드만 저장")
    parser.add_argument("--max-pages", type=int, default=45)
    parser.add_argument("--size", type=int, default=15)
    parser.add_argument("--delay", type=float, default=0.2)
    parser.add_argument("--limit", type=int, default=0, help="앞 n개 식당만 리뷰 수집. 테스트용")
    parser.add_argument("--scrolls", type=int, default=3)
    parser.add_argument("--fast-threshold", type=int, default=20)
    parser.add_argument("--max-scrolls", type=int, default=200)
    parser.add_argument("--target-coverage", type=float, default=0.95)
    parser.add_argument("--stall-limit", type=int, default=8)
    parser.add_argument("--stall-wait-ms", type=int, default=2000)
    parser.add_argument("--wait-ms", type=int, default=1000)
    parser.add_argument("--scroll-delay-ms", type=int, default=300)
    parser.add_argument("--headed", action="store_true", help="리뷰 수집 브라우저 창을 보이게 실행")
    parser.add_argument("--k", type=float, default=10.0, help="정량 순위표 리뷰 수 신뢰도 보정 상수")
    parser.add_argument("--weight-cap", type=float, default=10.0, help="리뷰어 가중치 상한")
    parser.add_argument("--penalty-scale", type=float, default=0.5, help="조작 의심 점수 감점 계수")
    parser.add_argument("--rating-weight", type=float, default=0.35, help="평균 별점 보너스 가중치")
    parser.add_argument("--rating-baseline", type=float, default=4.0, help="평균 별점 보너스 기준점")
    parser.add_argument("--plain-progress", action="store_true", help="고정형 화면 대신 기존 줄 단위 로그 출력")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    region = args.region or input("어느 지역을 검색할까요? 예: 인천 연수구 송도과학로\n> ").strip()
    if not region:
        print("지역명이 필요합니다.", file=sys.stderr)
        return 2
    if not args.api_key:
        print("KAKAO_REST_API_KEY 환경변수 또는 --api-key가 필요합니다.", file=sys.stderr)
        return 2
    if (
        any(ord(char) < 33 or ord(char) > 126 for char in args.api_key)
        or re.search(r"YOUR|발급|자신|REST_API|API_키", args.api_key, re.IGNORECASE)
    ):
        print(
            "KAKAO_REST_API_KEY에 한글 또는 안내용 예시 문구가 들어 있습니다. "
            "발급받은 실제 영문·숫자 REST API 키로 다시 설정하세요.",
            file=sys.stderr,
        )
        return 2
    run_id = f"{safe_name(region)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    run_dir = args.out_root.resolve() / run_id
    restaurants_dir = run_dir / "restaurants"
    raw_dir = run_dir / "raw_place_text"
    restaurants_dir.mkdir(parents=True, exist_ok=True)
    raw_dir.mkdir(parents=True, exist_ok=True)
    log_path = run_dir / "pipeline.log"
    dashboard = ConsoleDashboard(region, run_dir, plain=args.plain_progress)
    atexit.register(dashboard.close)
    dashboard.start()

    env = os.environ.copy()
    env["KAKAO_REST_API_KEY"] = args.api_key

    restaurant_search = [
        sys.executable,
        str(TOOLS / "kakao_restaurant_search.py"),
        region,
        "--mode",
        args.mode,
        "--radius",
        str(args.radius),
        "--outer-radius",
        str(args.outer_radius),
        "--grid-step",
        str(args.grid_step),
        "--cell-radius",
        str(args.cell_radius),
        "--max-pages",
        str(args.max_pages),
        "--size",
        str(args.size),
        "--delay",
        str(args.delay),
        "--out-dir",
        str(restaurants_dir),
    ]
    if args.minimal:
        restaurant_search.append("--minimal")
    run_step("1. 음식점 목록 수집", restaurant_search, stage=1, dashboard=dashboard, log_path=log_path, env=env)

    restaurant_xlsx = latest_xlsx(restaurants_dir, "kakao_restaurants_*.xlsx")
    combined_xlsx = run_dir / f"{safe_name(region)}_통합리뷰.xlsx"

    review_collect = [
        "node",
        str(TOOLS / "kakao_place_visible_text.mjs"),
        "--xlsx-file",
        str(restaurant_xlsx),
        "--out-dir",
        str(raw_dir),
        "--combined-xlsx",
        str(combined_xlsx),
        "--scrolls",
        str(args.scrolls),
        "--fast-threshold",
        str(args.fast_threshold),
        "--max-scrolls",
        str(args.max_scrolls),
        "--target-coverage",
        str(args.target_coverage),
        "--stall-limit",
        str(args.stall_limit),
        "--stall-wait-ms",
        str(args.stall_wait_ms),
        "--wait-ms",
        str(args.wait_ms),
        "--scroll-delay-ms",
        str(args.scroll_delay_ms),
    ]
    if args.limit:
        review_collect.extend(["--limit", str(args.limit)])
    if args.headed:
        review_collect.append("--headed")
    run_step("2. 카카오맵 리뷰 수집", review_collect, stage=2, dashboard=dashboard, log_path=log_path)

    if not combined_xlsx.exists():
        raise SystemExit(f"통합 리뷰 XLSX 생성 실패: {combined_xlsx}")

    ranking_html = run_dir / f"{safe_name(region)}_정량순위표.html"
    quantitative_rank = [
        sys.executable,
        str(TOOLS / "kakao_review_ranker.py"),
        str(combined_xlsx),
        "--out",
        str(ranking_html),
        "--k",
        str(args.k),
        "--weight-cap",
        str(args.weight_cap),
        "--penalty-scale",
        str(args.penalty_scale),
        "--rating-weight",
        str(args.rating_weight),
        "--rating-baseline",
        str(args.rating_baseline),
    ]
    run_step("3. 정량 순위표 생성", quantitative_rank, stage=3, dashboard=dashboard, log_path=log_path)
    dashboard.complete(
        run_dir,
        [combined_xlsx, ranking_html, ranking_html.with_suffix(".csv")],
    )
    dashboard.close()

    if not dashboard.enabled:
        print()
        print("완료")
        print(f"저장 위치: {run_dir}")
        print("XLSX, HTML, CSV를 생성했습니다.")
        print(
            "파일 이름: "
            + " / ".join(
                [
                    combined_xlsx.name,
                    ranking_html.name,
                    ranking_html.with_suffix(".csv").name,
                ]
            )
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
