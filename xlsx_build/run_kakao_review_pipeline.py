from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path


TOOLS = Path(__file__).resolve().parent
ROOT = TOOLS
DEFAULT_OUT_ROOT = ROOT / "outputs" / "review_pipeline"


def safe_name(value: str) -> str:
    text = "".join(ch if ch.isalnum() else "_" for ch in value.strip())
    return "_".join(part for part in text.split("_") if part) or "region"


def run_step(label: str, command: list[str], env: dict[str, str] | None = None) -> None:
    print(f"\n[{label}]")
    print(" ".join(f'"{part}"' if " " in part else part for part in command))
    result = subprocess.run(command, cwd=ROOT, env=env)
    if result.returncode != 0:
        raise SystemExit(f"{label} 실패: exit code {result.returncode}")


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

    run_id = f"{safe_name(region)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    run_dir = args.out_root.resolve() / run_id
    restaurants_dir = run_dir / "restaurants"
    raw_dir = run_dir / "raw_place_text"
    restaurants_dir.mkdir(parents=True, exist_ok=True)
    raw_dir.mkdir(parents=True, exist_ok=True)

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
    run_step("1. 음식점 목록 수집", restaurant_search, env=env)

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
    run_step("2. 카카오맵 리뷰 수집", review_collect)

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
    run_step("3. 정량 순위표 생성", quantitative_rank)

    print("\n완료")
    print(f"실행 폴더: {run_dir}")
    print(f"음식점 목록 XLSX: {restaurant_xlsx}")
    print(f"통합 리뷰 XLSX: {combined_xlsx}")
    print(f"정량 순위표 HTML: {ranking_html}")
    print(f"정량 순위표 CSV:  {ranking_html.with_suffix('.csv')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
