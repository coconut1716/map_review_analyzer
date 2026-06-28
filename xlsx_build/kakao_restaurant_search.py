from __future__ import annotations

import argparse
import csv
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.error import HTTPError
from urllib.request import Request, urlopen


API_BASE = "https://dapi.kakao.com"
ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUT_DIR = ROOT / "outputs" / "kakao_restaurants"


def request_json(path: str, params: dict[str, str | int], api_key: str) -> dict:
    url = f"{API_BASE}{path}?{urlencode(params)}"
    req = Request(
        url,
        headers={
            "Authorization": f"KakaoAK {api_key}",
            "Accept": "application/json",
            "User-Agent": "kakao-restaurant-search/1.0",
        },
    )
    try:
        with urlopen(req, timeout=20) as res:
            return json.loads(res.read().decode("utf-8"))
    except HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Kakao API 요청 실패: HTTP {exc.code}\nURL: {url}\n응답: {body}") from exc


def search_keyword(api_key: str, query: str, max_pages: int, size: int, delay: float) -> list[dict]:
    rows = []
    for page in range(1, max_pages + 1):
        payload = request_json(
            "/v2/local/search/keyword.json",
            {"query": query, "page": page, "size": size},
            api_key,
        )
        rows.extend(payload.get("documents", []))
        if payload.get("meta", {}).get("is_end", True):
            break
        time.sleep(delay)
    return rows


def resolve_center(api_key: str, region: str) -> tuple[str, str] | None:
    payload = request_json(
        "/v2/local/search/address.json",
        {"query": region, "size": 1},
        api_key,
    )
    docs = payload.get("documents", [])
    if docs:
        return docs[0]["x"], docs[0]["y"]

    payload = request_json(
        "/v2/local/search/keyword.json",
        {"query": region, "size": 1},
        api_key,
    )
    docs = payload.get("documents", [])
    if docs:
        return docs[0]["x"], docs[0]["y"]
    return None


def search_food_category(
    api_key: str,
    x: str,
    y: str,
    radius: int,
    max_pages: int,
    size: int,
    delay: float,
) -> list[dict]:
    rows = []
    for page in range(1, max_pages + 1):
        payload = request_json(
            "/v2/local/search/category.json",
            {
                "category_group_code": "FD6",
                "x": x,
                "y": y,
                "radius": radius,
                "page": page,
                "size": size,
                "sort": "distance",
            },
            api_key,
        )
        rows.extend(payload.get("documents", []))
        if payload.get("meta", {}).get("is_end", True):
            break
        time.sleep(delay)
    return rows


def normalize(doc: dict, region: str, source: str) -> dict:
    place_url = doc.get("place_url", "")
    review_url = f"{place_url.split('#', 1)[0]}#review" if place_url else ""
    return {
        "id": doc.get("id", ""),
        "place_name": doc.get("place_name", ""),
        "category_name": doc.get("category_name", ""),
        "category_group_name": doc.get("category_group_name", ""),
        "phone": doc.get("phone", ""),
        "address_name": doc.get("address_name", ""),
        "road_address_name": doc.get("road_address_name", ""),
        "x": doc.get("x", ""),
        "y": doc.get("y", ""),
        "distance": doc.get("distance", ""),
        "place_url": place_url,
        "review_url": review_url,
        "region_query": region,
        "source": source,
    }


def dedupe(rows: list[dict]) -> list[dict]:
    out = {}
    for row in rows:
        key = row.get("id") or f"{row.get('place_name')}|{row.get('address_name')}"
        out[key] = row
    return sorted(out.values(), key=lambda r: (r.get("distance") or "999999", r.get("place_name", "")))


def save_csv(path: Path, rows: list[dict], fields: list[str]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def save_xlsx(path: Path, rows: list[dict], fields: list[str]) -> bool:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        return False

    wb = Workbook()
    ws = wb.active
    ws.title = "restaurants"
    ws.append(fields)
    for row in rows:
        ws.append([row.get(field, "") for field in fields])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    ws.freeze_panes = "A2"

    widths = {
        "place_name": 26,
        "category_name": 34,
        "phone": 16,
        "address_name": 34,
        "road_address_name": 38,
        "place_url": 38,
        "review_url": 44,
    }
    for idx, field in enumerate(fields, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(field, 16)

    wb.save(path)
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="Kakao Local API로 특정 지역 음식점 목록과 리뷰 링크를 저장합니다.")
    parser.add_argument("region", help="검색할 지역. 예: '인천 연수구 송도동'")
    parser.add_argument("--api-key", default=os.getenv("KAKAO_REST_API_KEY"), help="Kakao REST API 키. 기본값: KAKAO_REST_API_KEY")
    parser.add_argument("--mode", choices=["keyword", "category", "both"], default="both")
    parser.add_argument("--radius", type=int, default=2000, help="category 모드 반경(m), 최대 20000")
    parser.add_argument("--max-pages", type=int, default=45)
    parser.add_argument("--size", type=int, default=15, help="페이지당 개수, 최대 15")
    parser.add_argument("--delay", type=float, default=0.2, help="페이지 요청 사이 대기초")
    parser.add_argument("--out-dir", type=Path, default=DEFAULT_OUT_DIR)
    args = parser.parse_args()

    if not args.api_key:
        print("KAKAO_REST_API_KEY 환경변수 또는 --api-key가 필요합니다.", file=sys.stderr)
        return 2
    if not (1 <= args.size <= 15):
        print("--size는 1~15 사이여야 합니다.", file=sys.stderr)
        return 2
    if not (1 <= args.radius <= 20000):
        print("--radius는 1~20000 사이여야 합니다.", file=sys.stderr)
        return 2

    collected = []
    if args.mode in ("keyword", "both"):
        query = f"{args.region} 음식점"
        collected.extend(normalize(doc, args.region, "keyword") for doc in search_keyword(args.api_key, query, args.max_pages, args.size, args.delay))

    if args.mode in ("category", "both"):
        center = resolve_center(args.api_key, args.region)
        if center:
            x, y = center
            collected.extend(
                normalize(doc, args.region, "category")
                for doc in search_food_category(args.api_key, x, y, args.radius, args.max_pages, args.size, args.delay)
            )
        else:
            print(f"지역 중심 좌표를 찾지 못해 category 검색을 건너뜁니다: {args.region}", file=sys.stderr)

    rows = dedupe(collected)
    now = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    safe_region = "".join(ch if ch.isalnum() else "_" for ch in args.region).strip("_")
    out_dir = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    base = out_dir / f"kakao_restaurants_{safe_region}_{now}"

    fields = [
        "id",
        "place_name",
        "category_name",
        "category_group_name",
        "phone",
        "address_name",
        "road_address_name",
        "x",
        "y",
        "distance",
        "place_url",
        "review_url",
        "region_query",
        "source",
    ]
    json_path = base.with_suffix(".json")
    csv_path = base.with_suffix(".csv")
    xlsx_path = base.with_suffix(".xlsx")

    json_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    save_csv(csv_path, rows, fields)
    xlsx_ok = save_xlsx(xlsx_path, rows, fields)

    print(f"수집 식당 수: {len(rows)}")
    print(f"JSON: {json_path}")
    print(f"CSV:  {csv_path}")
    if xlsx_ok:
        print(f"XLSX: {xlsx_path}")
    else:
        print("XLSX: openpyxl이 없어 생략했습니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

