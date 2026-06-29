from __future__ import annotations

import argparse
import csv
import html
import math
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path


NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def col_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall("m:si", NS):
        strings.append("".join((t.text or "") for t in si.findall(".//m:t", NS)))
    return strings


def resolve_sheet_targets(zf: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    targets = {}
    for sheet in workbook.findall(".//m:sheet", NS):
        name = sheet.attrib["name"]
        rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = relmap[rid]
        if target.startswith("/"):
            target = target.lstrip("/")
        elif not target.startswith("xl/"):
            target = "xl/" + target
        targets[name] = target
    return targets


def read_sheet(zf: zipfile.ZipFile, target: str, shared_strings: list[str]) -> list[list[str]]:
    root = ET.fromstring(zf.read(target))
    rows = []
    for row in root.findall(".//m:sheetData/m:row", NS):
        values = {}
        max_col = -1
        for cell in row.findall("m:c", NS):
            idx = col_index(cell.attrib.get("r", ""))
            max_col = max(max_col, idx)
            value = ""
            value_node = cell.find("m:v", NS)
            if value_node is not None:
                raw = value_node.text or ""
                if cell.attrib.get("t") == "s" and raw.isdigit():
                    value = shared_strings[int(raw)]
                else:
                    value = raw
            elif cell.attrib.get("t") == "inlineStr":
                value = "".join((t.text or "") for t in cell.findall(".//m:t", NS))
            values[idx] = value
        if max_col >= 0:
            rows.append([values.get(i, "") for i in range(max_col + 1)])
    return rows


def read_workbook(path: Path) -> dict[str, list[list[str]]]:
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        targets = resolve_sheet_targets(zf)
        return {name: read_sheet(zf, target, shared_strings) for name, target in targets.items()}


def rows_as_dicts(rows: list[list[str]]) -> list[dict[str, str]]:
    if not rows:
        return []
    headers = [str(cell).strip() for cell in rows[0]]
    out = []
    for row in rows[1:]:
        item = {}
        for i, header in enumerate(headers):
            if header:
                item[header] = row[i] if i < len(row) else ""
        if any(str(v).strip() for v in item.values()):
            out.append(item)
    return out


def to_float(value, default: float | None = None) -> float | None:
    if value is None:
        return default
    text = str(value).strip().replace(",", "")
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def to_int(value, default: int = 0) -> int:
    number = to_float(value, None)
    return default if number is None else int(round(number))


def safe_score(value: float | None) -> str:
    return "" if value is None else f"{value:.4f}"


def clamp(value: float, low: float = 0.0, high: float = 1.0) -> float:
    return max(low, min(high, value))


def parse_date(value: str):
    text = str(value or "").strip().rstrip(".")
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def max_rolling_count(dates, days: int) -> int:
    if not dates:
        return 0
    ordered = sorted(dates)
    best = 0
    left = 0
    for right, date in enumerate(ordered):
        while (date - ordered[left]).days > days:
            left += 1
        best = max(best, right - left + 1)
    return best


def repeat_burst_risk(dates, n: int) -> float:
    if n < 10 or not dates:
        return 0.0
    ordered = sorted(dates)
    clusters = []
    current = [ordered[0]]
    for date in ordered[1:]:
        if (date - current[-1]).days <= 7:
            current.append(date)
        else:
            clusters.append(current)
            current = [date]
    clusters.append(current)

    burst_min = max(4, math.ceil(n * 0.2))
    burst_clusters = [cluster for cluster in clusters if len(cluster) >= burst_min]
    if len(burst_clusters) < 2:
        return 0.0
    for i in range(len(burst_clusters) - 1):
        gap = (burst_clusters[i + 1][0] - burst_clusters[i][-1]).days
        if gap >= 30:
            return 1.0
    return 0.5


def reviewer_count_weight(count: int) -> float:
    n = max(0, count)
    if n <= 10:
        return 1 + n / 3
    if n <= 100:
        return 1 + 10 / 3 + 1.2 * math.log(n / 10)
    return 1 + 10 / 3 + 1.2 * math.log(10) + 0.7 * math.log10(n / 100)


def follower_weight(count: int) -> float:
    return reviewer_count_weight(count)


def reviewer_weight(review_count: int, follower_count: int, weight_cap: float) -> float:
    weight = reviewer_count_weight(review_count) * follower_weight(follower_count)
    return min(weight_cap, weight) if weight_cap > 0 else weight


def store_rating_effect(rating: float | None, baseline: float) -> float:
    if rating is None:
        return 0.0
    return clamp(rating - baseline, -1.0, 1.0)

def build_burst_clusters(dates, max_gap_days: int = 7) -> list[list]:
    if not dates:
        return []
    ordered = sorted(dates)
    clusters = []
    current = [ordered[0]]
    for date in ordered[1:]:
        if (date - current[-1]).days <= max_gap_days:
            current.append(date)
        else:
            clusters.append(current)
            current = [date]
    clusters.append(current)
    return clusters


def cluster_center_ordinal(cluster) -> float:
    return sum(date.toordinal() for date in cluster) / len(cluster)


def coefficient_of_variation(values: list[float]) -> float | None:
    values = [value for value in values if value is not None]
    if len(values) < 2:
        return None
    mean = sum(values) / len(values)
    if mean <= 0:
        return None
    variance = sum((value - mean) ** 2 for value in values) / len(values)
    return math.sqrt(variance) / mean


def manipulation_metrics(valid: list[dict]) -> dict[str, float | int]:
    n = len(valid)
    empty = {
        "날짜집중위험": 0.0,
        "동일일자위험": 0.0,
        "버스트리뷰비율": 0.0,
        "버스트밀도위험": 0.0,
        "버스트클러스터수": 0,
        "버스트주기성위험": 0.0,
        "리뷰간격균일위험": 0.0,
        "보정리뷰간격균일위험": 0.0,
        "저활동고평점위험": 0.0,
        "만점성향위험": 0.0,
        "반복버스트위험": 0.0,
        "조작의심점수": 0.0,
        "최대7일리뷰수": 0,
        "최대3일리뷰수": 0,
        "최대동일일자리뷰수": 0,
        "별5유효리뷰수": 0,
    }
    if n == 0:
        return empty

    five_star_valid = [item for item in valid if item["given"] >= 5.0 and item.get("date")]
    date_n = len(five_star_valid)
    dates = [item["date"] for item in five_star_valid]
    max7 = max_rolling_count(dates, 7)
    max3 = max_rolling_count(dates, 3)
    by_date = defaultdict(int)
    for date in dates:
        by_date[date] += 1
    same_day_max = max(by_date.values(), default=0)

    burst_ratio = max7 / date_n if date_n else 0.0
    same_day_ratio = same_day_max / date_n if date_n else 0.0
    low_activity_high = sum(
        1
        for item in valid
        if item["given"] >= 4.5 and item["reviewer_count"] <= 3 and item["follower_count"] == 0
    ) / n
    perfect_bias = sum(1 for item in valid if item["given"] >= 5.0 and item["reviewer_avg"] >= 4.8) / n

    date_burst_risk = clamp((burst_ratio - 0.18) / 0.32) if date_n >= 5 else 0.0
    same_day_risk = clamp((same_day_ratio - 0.12) / 0.25) if date_n >= 5 else 0.0
    low_activity_risk = clamp((low_activity_high - 0.25) / 0.50)
    perfect_bias_risk = clamp((perfect_bias - 0.40) / 0.45)

    clusters = build_burst_clusters(dates, 7)
    burst_clusters = [cluster for cluster in clusters if len(cluster) >= 3]
    burst_review_count = sum(len(cluster) for cluster in burst_clusters)
    burst_density = burst_review_count / date_n if date_n else 0.0
    burst_density_risk = clamp((burst_density - 0.25) / 0.45) if date_n >= 8 else 0.0

    centers = [cluster_center_ordinal(cluster) for cluster in burst_clusters]
    center_gaps = [centers[i + 1] - centers[i] for i in range(len(centers) - 1)]
    center_cv = coefficient_of_variation(center_gaps)
    periodic_risk = clamp((0.35 - center_cv) / 0.35) if center_cv is not None and len(burst_clusters) >= 3 else 0.0

    unique_dates = sorted(set(dates))
    day_gaps = [(unique_dates[i] - unique_dates[i - 1]).days for i in range(1, len(unique_dates))]
    day_gap_cv = coefficient_of_variation(day_gaps)
    regular_risk = clamp((0.45 - day_gap_cv) / 0.45) if day_gap_cv is not None and date_n >= 8 else 0.0
    regular_adjusted_risk = regular_risk * burst_density_risk

    repeated_burst_risk = 1.0 if len(burst_clusters) >= 2 and any(
        (burst_clusters[i + 1][0] - burst_clusters[i][-1]).days >= 30 for i in range(len(burst_clusters) - 1)
    ) else 0.5 if len(burst_clusters) >= 2 else 0.0
    risk = clamp(
        0.25 * date_burst_risk
        + 0.15 * same_day_risk
        + 0.25 * burst_density_risk
        + 0.15 * periodic_risk
        + 0.10 * regular_adjusted_risk
        + 0.06 * low_activity_risk
        + 0.04 * perfect_bias_risk
    )

    return {
        "날짜집중위험": date_burst_risk,
        "동일일자위험": same_day_risk,
        "버스트리뷰비율": burst_density,
        "버스트밀도위험": burst_density_risk,
        "버스트클러스터수": len(burst_clusters),
        "버스트주기성위험": periodic_risk,
        "리뷰간격균일위험": regular_risk,
        "보정리뷰간격균일위험": regular_adjusted_risk,
        "저활동고평점위험": low_activity_risk,
        "만점성향위험": perfect_bias_risk,
        "반복버스트위험": repeated_burst_risk,
        "조작의심점수": risk,
        "최대7일리뷰수": max7,
        "최대3일리뷰수": max3,
        "최대동일일자리뷰수": same_day_max,
        "별5유효리뷰수": date_n,
    }


def compute_rankings(places: list[dict[str, str]], reviews: list[dict[str, str]], k: float, weight_cap: float, penalty_scale: float, rating_weight: float, rating_baseline: float) -> list[dict]:
    place_by_id = {str(p.get("식당번호", "")).strip(): p for p in places}
    grouped = defaultdict(list)
    for review in reviews:
        grouped[str(review.get("식당번호", "")).strip()].append(review)

    ranked = []
    for place_id, place in place_by_id.items():
        valid = []
        for review in grouped.get(place_id, []):
            given = to_float(review.get("이 식당에 준 별점"), None)
            reviewer_avg = to_float(review.get("리뷰어 평균 별점"), None)
            if given is None or reviewer_avg is None:
                continue
            reviewer_count = max(0, to_int(review.get("리뷰어 총 리뷰수"), 0))
            follower_count = max(0, to_int(review.get("팔로워 수"), 0))
            weight = reviewer_weight(reviewer_count, follower_count, weight_cap)
            valid.append(
                {
                    "delta": given - reviewer_avg,
                    "weight": weight,
                    "given": given,
                    "reviewer_avg": reviewer_avg,
                    "reviewer_count": reviewer_count,
                    "follower_count": follower_count,
                    "date": parse_date(review.get("리뷰 날짜")),
                }
            )

        weight_sum = sum(item["weight"] for item in valid)
        raw = sum(item["delta"] * item["weight"] for item in valid) / weight_sum if weight_sum else None
        n = len(valid)
        confidence = n / (n + k) if n else 0
        before_risk = raw * confidence if raw is not None else None
        risk = manipulation_metrics(valid)
        manipulation_penalty = penalty_scale * risk["조작의심점수"]
        risk_adjusted = before_risk - manipulation_penalty if before_risk is not None else None
        store_rating = to_float(place.get("식당 평균 별점"), None)
        rating_effect = store_rating_effect(store_rating, rating_baseline)
        rating_trust = 1 - risk["조작의심점수"]
        rating_bonus = rating_weight * rating_trust * rating_effect if risk_adjusted is not None else 0.0
        adjusted = risk_adjusted + rating_bonus if risk_adjusted is not None else None
        quality = store_rating + adjusted if store_rating is not None and adjusted is not None else adjusted

        ranked.append(
            {
                "식당번호": place_id,
                "장소명": place.get("장소명", ""),
                "카테고리": place.get("카테고리", ""),
                "식당평균별점": store_rating,
                "식당총리뷰수": to_int(place.get("식당 총 리뷰수"), 0),
                "캡처리뷰수": to_int(place.get("캡처된 리뷰수"), 0),
                "유효리뷰수": n,
                "리뷰어평균대비점수": raw,
                "신뢰도보정": confidence,
                "조작전점수": before_risk,
                "조작의심점수": risk["조작의심점수"],
                "조작감점": manipulation_penalty,
                "조작후점수": risk_adjusted,
                "평균별점효과": rating_effect,
                "평균별점반영도": rating_trust,
                "평균별점보너스": rating_bonus,
                "최종보정점수": adjusted,
                "별점포함점수": quality,
                "평균리뷰별점": (sum(item["given"] for item in valid) / n) if n else None,
                "평균리뷰어평균": (sum(item["reviewer_avg"] for item in valid) / n) if n else None,
                "가중치합": weight_sum,
                "최대7일리뷰수": risk["최대7일리뷰수"],
                "최대3일리뷰수": risk["최대3일리뷰수"],
                "최대동일일자리뷰수": risk["최대동일일자리뷰수"],
                "별5유효리뷰수": risk["별5유효리뷰수"],
                "날짜집중위험": risk["날짜집중위험"],
                "동일일자위험": risk["동일일자위험"],
                "버스트리뷰비율": risk["버스트리뷰비율"],
                "버스트밀도위험": risk["버스트밀도위험"],
                "버스트클러스터수": risk["버스트클러스터수"],
                "버스트주기성위험": risk["버스트주기성위험"],
                "리뷰간격균일위험": risk["리뷰간격균일위험"],
                "보정리뷰간격균일위험": risk["보정리뷰간격균일위험"],
                "저활동고평점위험": risk["저활동고평점위험"],
                "만점성향위험": risk["만점성향위험"],
                "반복버스트위험": risk["반복버스트위험"],
                "원본URL": place.get("원본 URL", ""),
            }
        )

    ranked.sort(
        key=lambda row: (
            row["최종보정점수"] is not None,
            row["최종보정점수"] if row["최종보정점수"] is not None else -999,
            row["식당평균별점"] if row["식당평균별점"] is not None else -999,
            row["유효리뷰수"],
        ),
        reverse=True,
    )
    for i, row in enumerate(ranked, start=1):
        row["순위"] = i
    return ranked


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(input_path.stem + "_순위표.html")


def region_label_from_source(source: Path) -> str:
    stem = source.stem
    for suffix in ("_통합리뷰", "_통합리뷰_자동스크롤", "_combined_reviews", "_reviews"):
        if stem.endswith(suffix):
            stem = stem[: -len(suffix)]
    stem = stem.replace("kakao_restaurants_", "")
    parts = [part for part in stem.split("_") if part]
    if len(parts) >= 2 and parts[-2].isdigit() and parts[-1].isdigit():
        parts = parts[:-2]
    if len(parts) >= 1 and len(parts[-1]) == 8 and parts[-1].isdigit():
        parts = parts[:-1]
    if not parts:
        parent = source.parent.name
        parts = [part for part in parent.split("_") if part]
        if len(parts) >= 2 and parts[-2].isdigit() and parts[-1].isdigit():
            parts = parts[:-2]
    return " ".join(parts) if parts else source.stem
def write_csv(path: Path, rows: list[dict]) -> None:
    fields = [
        "순위",
        "장소명",
        "카테고리",
        "최종보정점수",
        "조작전점수",
        "조작의심점수",
        "조작감점",
        "조작후점수",
        "평균별점효과",
        "평균별점반영도",
        "평균별점보너스",
        "리뷰어평균대비점수",
        "신뢰도보정",
        "별점포함점수",
        "식당평균별점",
        "식당총리뷰수",
        "캡처리뷰수",
        "유효리뷰수",
        "평균리뷰별점",
        "평균리뷰어평균",
        "가중치합",
        "최대7일리뷰수",
        "최대3일리뷰수",
        "최대동일일자리뷰수",
        "별5유효리뷰수",
        "날짜집중위험",
        "동일일자위험",
        "버스트리뷰비율",
        "버스트밀도위험",
        "버스트클러스터수",
        "버스트주기성위험",
        "리뷰간격균일위험",
        "보정리뷰간격균일위험",
        "저활동고평점위험",
        "만점성향위험",
        "반복버스트위험",
        "원본URL",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: safe_score(row[field]) if isinstance(row.get(field), float) else row.get(field, "") for field in fields})


def write_html(path: Path, rows: list[dict], source: Path, k: float, weight_cap: float, penalty_scale: float, rating_weight: float, rating_baseline: float) -> None:
    region_label = region_label_from_source(source)
    top_count = sum(1 for row in rows if row["유효리뷰수"] > 0)
    table_rows = []
    for row in rows:
        score_class = "positive" if (row["최종보정점수"] or 0) > 0 else "negative" if row["최종보정점수"] is not None else "empty"
        risk_class = "risk-high" if row["조작의심점수"] >= 0.55 else "risk-mid" if row["조작의심점수"] >= 0.25 else "risk-low"
        table_rows.append(
            "<tr>"
            f"<td class='rank'>{row['순위']}</td>"
            f"<td class='name'><a href='{html.escape(row['원본URL'])}'>{html.escape(row['장소명'])}</a><span>{html.escape(row['카테고리'])}</span></td>"
            f"<td class='{score_class}'>{safe_score(row['최종보정점수'])}</td>"
            f"<td>{safe_score(row['조작전점수'])}</td>"
            f"<td class='{risk_class}'>{safe_score(row['조작의심점수'])}</td>"
            f"<td>{safe_score(row['조작감점'])}</td>"
            f"<td>{safe_score(row['평균별점보너스'])}</td>"
            f"<td>{safe_score(row['조작후점수'])}</td>"
            f"<td>{safe_score(row['리뷰어평균대비점수'])}</td>"
            f"<td>{safe_score(row['신뢰도보정'])}</td>"
            f"<td>{safe_score(row['식당평균별점'])}</td>"
            f"<td>{row['식당총리뷰수']}</td>"
            f"<td>{row['캡처리뷰수']}</td>"
            f"<td>{row['유효리뷰수']}</td>"
            f"<td>{row['최대7일리뷰수']}</td>"
            f"<td>{row['최대동일일자리뷰수']}</td>"
            f"<td>{row['별5유효리뷰수']}</td>"
            f"<td>{row['버스트클러스터수']}</td>"
            f"<td>{safe_score(row['버스트리뷰비율'])}</td>"
            f"<td>{safe_score(row['버스트밀도위험'])}</td>"
            f"<td>{safe_score(row['버스트주기성위험'])}</td>"
            f"<td>{safe_score(row['보정리뷰간격균일위험'])}</td>"
            f"<td>{safe_score(row['저활동고평점위험'])}</td>"
            "</tr>"
        )

    content = fr"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>카카오 리뷰 보정 순위표</title>
<style>
  body {{ margin: 0; background: #f7f8fb; color: #1f2937; font-family: "Malgun Gothic", Arial, sans-serif; }}
  main {{ width: min(1680px, calc(100vw - 32px)); margin: 28px auto 56px; }}
  h1 {{ margin: 0 0 8px; font-size: 28px; }}
  .meta {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin: 18px 0 18px; }}
  .card {{ background: white; border: 1px solid #e5e7eb; border-radius: 8px; padding: 12px 14px; }}
  .card b {{ display: block; font-size: 20px; margin-top: 4px; }}
  .formula {{ background: #eef6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 16px 18px; line-height: 1.7; margin-bottom: 18px; }}
  .formula h2 {{ margin: 0 0 12px; font-size: 18px; }}
  .formula p {{ margin: 10px 0; }}
  .formula .formula-note {{ color: #374151; font-size: 13px; }}
  .formula mjx-container[display="true"] {{ margin: 8px 0 14px !important; }}
  table {{ width: 100%; border-collapse: collapse; background: white; border: 1px solid #d1d5db; table-layout: fixed; }}
  th, td {{ border-bottom: 1px solid #e5e7eb; padding: 9px 8px; text-align: right; vertical-align: top; font-size: 12px; }}
  th {{ position: sticky; top: 0; background: #1f4e78; color: white; z-index: 1; }}
  td.rank {{ text-align: center; font-weight: 800; width: 48px; }}
  td.name, th.name {{ text-align: left; width: 220px; }}
  td.name a {{ color: #111827; font-weight: 800; text-decoration: none; }}
  td.name span {{ display: block; margin-top: 3px; color: #6b7280; font-size: 12px; }}
  .positive {{ color: #047857; font-weight: 800; }}
  .negative {{ color: #b91c1c; font-weight: 800; }}
  .empty {{ color: #9ca3af; }}
  .risk-low {{ color: #047857; font-weight: 800; }}
  .risk-mid {{ color: #b45309; font-weight: 800; }}
  .risk-high {{ color: #b91c1c; font-weight: 900; }}
  .note {{ color: #6b7280; font-size: 13px; margin-top: 10px; }}
  @media (max-width: 900px) {{
    main {{ width: 100%; margin: 0; padding: 14px; overflow-x: auto; }}
    .meta {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
    table {{ min-width: 1760px; }}
  }}
</style>
<script>
  window.MathJax = {{
    tex: {{
      inlineMath: [["\\(", "\\)"]],
      displayMath: [["\\[", "\\]"]]
    }},
    svg: {{ fontCache: "global" }}
  }};
</script>
<script defer src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js"></script>
</head>
<body>
<main>
  <h1>카카오 리뷰 보정 순위표</h1>
  <div class="note">지역: {html.escape(region_label)}</div>
  <section class="meta">
    <div class="card">식당 수<b>{len(rows)}</b></div>
    <div class="card">유효 리뷰가 있는 식당<b>{top_count}</b></div>
    <div class="card">리뷰 수 보정 k<b>{k:g}</b></div>
    <div class="card">조작 감점 계수<b>{penalty_scale:g}</b></div>
    <div class="card">평균별점 기준<b>{rating_baseline:g}</b></div>
    <div class="card">평균별점 가중치<b>{rating_weight:g}</b></div>
  </section>
  <section class="formula">
    <h2>순위식</h2>
    <p>리뷰별 차이는 해당 식당에 준 별점에서 리뷰어의 평소 평균 별점을 뺀 값입니다.</p>
    \[
    d_{{ij}} = r_{{ij}} - \bar{{r}}_i
    \]
    <p>리뷰 수 가중치 \(\phi(n)\)는 리뷰어 총 리뷰 수가 많을수록 커지되, 구간별로 증가 속도를 줄이는 누진식입니다.</p>
    \[
    \phi(n)=
    \begin{{cases}}
    1+\dfrac{{n}}{{3}}, & n\le 10 \\
    1+\dfrac{{10}}{{3}}+1.2\ln\left(\dfrac{{n}}{{10}}\right), & 10<n\le 100 \\
    1+\dfrac{{10}}{{3}}+1.2\ln(10)+0.7\log_{{10}}\left(\dfrac{{n}}{{100}}\right), & n>100
    \end{{cases}}
    \]
    <p>팔로워 가중치도 리뷰 수 가중치와 같은 누진식을 사용합니다.</p>
    \[
    \psi(f_i)=\phi(f_i)
    \]
    <p>최종 리뷰어 가중치는 리뷰 수 가중치와 팔로워 가중치를 곱한 뒤 상한을 적용합니다.</p>
    \[
    w_i=\min\left({weight_cap:g},\phi(n_i)\psi(f_i)\right)
    \]
    <p>조작 전 점수는 리뷰별 차이의 가중 평균에 유효 리뷰 수 신뢰도 보정을 곱한 값입니다.</p>
    \[
    B_j=
    \frac{{\sum_{{i\in I_j}} d_{{ij}}w_i}}{{\sum_{{i\in I_j}}w_i}}
    \times
    \frac{{m_j}}{{m_j+{k:g}}}
    \]
    <p>조작의심점수는 날짜 집중, 동일 일자 집중, 버스트 밀도, 주기성, 간격 균일성, 저활동 고평점, 만점 성향을 합산합니다.</p>
    \[
    G_j=0.25A_j+0.15D_j+0.25B^{{density}}_j+0.15P_j+0.10R^{{regular*}}_j+0.06U_j+0.04M_j
    \]
    <p class="formula-note">날짜/버스트/주기성 위험은 별 5개 리뷰만 대상으로 봅니다. 날짜집중은 별 5개 리뷰 중 가장 빽빽한 7일 구간, 버스트밀도는 별 5개 리뷰 3개 이상이 7일 간격으로 붙은 묶음의 비율입니다.</p>
    <p>조작 후 점수는 조작 전 점수에서 조작의심점수 기반 감점을 뺀 값입니다.</p>
    \[
    Q_j=B_j-({penalty_scale:g}\times G_j)
    \]
    <p>평균별점보너스는 조작 의심이 낮을수록 식당 평균 별점을 더 신뢰하도록 설계했습니다.</p>
    \[
    R_j={rating_weight:g}\times(1-G_j)\times\operatorname{{clip}}(a_j-{rating_baseline:g},-1,1)
    \]
    <p>최종보정점수는 조작 후 점수와 평균별점보너스를 더한 값입니다.</p>
    \[
    F_j=Q_j+R_j
    \]
  </section>
  <table>
    <thead>
      <tr>
        <th>순위</th><th class="name">식당</th><th>최종</th><th>조작전</th><th>의심</th><th>감점</th><th>별점보너스</th><th>조작후</th><th>원점수</th><th>신뢰도</th><th>식당평균</th><th>총리뷰</th><th>캡처</th><th>유효</th><th>최대7일</th><th>동일일</th><th>별5유효</th><th>버스트묶음</th><th>버스트비율</th><th>버스트밀도</th><th>주기성</th><th>균일보정</th><th>저활동고평점</th>
      </tr>
    </thead>
    <tbody>
      {''.join(table_rows)}
    </tbody>
  </table>
</main>
</body>
</html>
"""
    path.write_text(content, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="카카오 통합 리뷰 XLSX에서 보정 순위표 HTML/CSV를 생성합니다.")
    parser.add_argument("xlsx_path", type=Path, help="송도과학로_통합리뷰.xlsx 같은 통합 리뷰 파일 경로")
    parser.add_argument("--out", type=Path, default=None, help="출력 HTML 경로. 기본값: 입력파일명_순위표.html")
    parser.add_argument("--k", type=float, default=10.0, help="리뷰 수 신뢰도 보정 상수")
    parser.add_argument("--weight-cap", type=float, default=10.0, help="리뷰어 가중치 상한. 0이면 상한 없음")
    parser.add_argument("--penalty-scale", type=float, default=0.5, help="조작 의심 점수 감점 계수")
    parser.add_argument("--rating-weight", type=float, default=0.35, help="평균 별점 보너스 가중치")
    parser.add_argument("--rating-baseline", type=float, default=4.0, help="평균 별점 보너스 기준점")
    args = parser.parse_args()

    if not args.xlsx_path.exists():
        print(f"파일을 찾지 못했습니다: {args.xlsx_path}", file=sys.stderr)
        return 2

    workbook = read_workbook(args.xlsx_path)
    if "식당요약" not in workbook or "리뷰전체" not in workbook:
        print("필수 시트가 없습니다. '식당요약'과 '리뷰전체' 시트가 필요합니다.", file=sys.stderr)
        return 2

    places = rows_as_dicts(workbook["식당요약"])
    reviews = rows_as_dicts(workbook["리뷰전체"])
    rankings = compute_rankings(places, reviews, args.k, args.weight_cap, args.penalty_scale, args.rating_weight, args.rating_baseline)

    html_path = args.out or default_output_path(args.xlsx_path)
    html_path.parent.mkdir(parents=True, exist_ok=True)
    csv_path = html_path.with_suffix(".csv")
    write_html(html_path, rankings, args.xlsx_path, args.k, args.weight_cap, args.penalty_scale, args.rating_weight, args.rating_baseline)
    write_csv(csv_path, rankings)

    print(f"HTML: {html_path}")
    print(f"CSV:  {csv_path}")
    print(f"식당 수: {len(rankings)}, 리뷰 있는 식당: {sum(1 for row in rankings if row['유효리뷰수'] > 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())













