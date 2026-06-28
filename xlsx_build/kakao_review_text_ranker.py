from __future__ import annotations

import argparse
import csv
import html
import re
import sys
from collections import defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from kakao_review_ranker import (  # noqa: E402
    compute_rankings,
    read_workbook,
    rows_as_dicts,
    safe_score,
    to_float,
)

DIRECT_MANIP_PATTERNS = [
    r"리뷰\s*알바",
    r"별점\s*알바",
    r"평점\s*알바",
    r"알바\s*리뷰",
    r"가짜\s*리뷰",
    r"리뷰\s*조작",
    r"별점\s*조작",
    r"평점\s*조작",
    r"조작\s*리뷰",
    r"리뷰.*?수상",
    r"리뷰.*?의심",
    r"알바.*?의심",
]

PROMO_PATTERNS = [
    r"체험단",
    r"협찬",
    r"제공\s*받",
    r"원고료",
    r"리뷰\s*이벤트(?!\s*(말고|아님|아니|가\s*아니|때문에\s*쓰는\s*게\s*아니))",
    r"리뷰이벤트(?!\s*(말고|아님|아니|가\s*아니|때문에\s*쓰는\s*게\s*아니))",
    r"평점.{0,12}리뷰\s*이벤트",
    r"리뷰\s*이벤트.{0,12}(뻥튀기|믿지|속은|점수|평점)",
    r"광고\s*(ㄹㅈㄷ|레전드|같|티|리뷰|빨|느낌)",
]

SEVERE_NEG_PATTERNS = [
    r"최악",
    r"다신\s*(안|못|는|가지|안가)",
    r"재방문\s*(안|없)",
    r"비추",
    r"가지\s*마",
    r"가지마",
    r"맛\s*없",
    r"맛없",
    r"노맛",
    r"불친절",
    r"위생",
    r"머리카락",
    r"벌레",
    r"곰팡",
    r"상했",
    r"식중독",
    r"배탈",
    r"토할",
    r"환불",
    r"컴플레인",
]

NEG_PATTERNS = [
    r"별로",
    r"실망",
    r"아쉽",
    r"느끼",
    r"짜다",
    r"짬",
    r"싱겁",
    r"비싸",
    r"바가지",
    r"늦",
    r"오래\s*걸",
    r"기다",
    r"딱딱",
    r"질김",
    r"퍽퍽",
    r"냄새",
    r"시끄",
    r"불편",
    r"차갑",
    r"부족",
    r"작다",
    r"적다",
]

POS_PATTERNS = [
    r"맛있",
    r"맛집",
    r"친절",
    r"추천",
    r"재방문",
    r"또\s*갈",
    r"또갈",
    r"깔끔",
    r"신선",
    r"만족",
    r"최고",
    r"좋",
    r"가성비",
    r"푸짐",
    r"든든",
    r"부드럽",
    r"고소",
]

COMPILED = {
    "direct": [re.compile(p) for p in DIRECT_MANIP_PATTERNS],
    "promo": [re.compile(p) for p in PROMO_PATTERNS],
    "severe": [re.compile(p) for p in SEVERE_NEG_PATTERNS],
    "negative": [re.compile(p) for p in NEG_PATTERNS],
    "positive": [re.compile(p) for p in POS_PATTERNS],
}


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "").strip().lower())


def count_matches(patterns, text: str) -> int:
    return sum(1 for pattern in patterns if pattern.search(text))


def snippet(text: str, patterns) -> str:
    raw = str(text or "").replace("\n", " ").strip()
    lowered = raw.lower()
    for pattern in patterns:
        m = pattern.search(lowered)
        if m:
            start = max(0, m.start() - 26)
            end = min(len(raw), m.end() + 46)
            return raw[start:end]
    return raw[:90]


def analyze_reviews(reviews: list[dict[str, str]]) -> dict:
    text_reviews = []
    direct = promo = severe = negative = positive = 0
    red_snippets = []
    neg_snippets = []

    for review in reviews:
        text = norm(review.get("리뷰 내용", ""))
        if not text:
            continue
        text_reviews.append(review)
        d = count_matches(COMPILED["direct"], text)
        p = count_matches(COMPILED["promo"], text)
        s = count_matches(COMPILED["severe"], text)
        n = count_matches(COMPILED["negative"], text)
        po = count_matches(COMPILED["positive"], text)
        direct += 1 if d else 0
        promo += 1 if p else 0
        severe += 1 if s else 0
        negative += 1 if n else 0
        positive += 1 if po else 0
        if d or p:
            red_snippets.append(snippet(review.get("리뷰 내용", ""), COMPILED["direct"] + COMPILED["promo"]))
        elif s and len(neg_snippets) < 5:
            neg_snippets.append(snippet(review.get("리뷰 내용", ""), COMPILED["severe"]))

    m = len(text_reviews)
    if m == 0:
        return {
            "텍스트리뷰수": 0,
            "직접조작언급": 0,
            "광고협찬언급": 0,
            "강한불만리뷰": 0,
            "일반불만리뷰": 0,
            "긍정리뷰": 0,
            "텍스트감점": 0.0,
            "텍스트보너스": 0.0,
            "텍스트신호점수": 0.0,
            "텍스트요약": "리뷰 텍스트 없음",
            "대표문구": "",
        }

    direct_ratio = direct / m
    promo_ratio = promo / m
    severe_ratio = severe / m
    neg_ratio = negative / m
    pos_ratio = positive / m

    direct_penalty = min(0.65, (0.22 if direct else 0.0) + 0.85 * direct_ratio)
    promo_penalty = min(0.28, 0.35 * promo_ratio)
    severe_penalty = min(0.45, 0.62 * severe_ratio)
    neg_penalty = min(0.22, 0.20 * neg_ratio)
    text_penalty = min(0.95, direct_penalty + promo_penalty + severe_penalty + neg_penalty)

    positive_excess = max(0.0, pos_ratio - severe_ratio - 0.5 * neg_ratio)
    text_confidence = m / (m + 10)
    text_bonus = min(0.06, 0.05 * positive_excess * text_confidence)
    signal = text_bonus - text_penalty

    summary_bits = []
    if direct:
        summary_bits.append(f"조작/알바 직접언급 {direct}건")
    if promo:
        summary_bits.append(f"광고/협찬/이벤트 언급 {promo}건")
    if severe:
        summary_bits.append(f"강한 불만 {severe}건")
    if negative:
        summary_bits.append(f"일반 불만 {negative}건")
    if positive:
        summary_bits.append(f"긍정 신호 {positive}건")
    summary = ", ".join(summary_bits) if summary_bits else "특이 텍스트 신호 적음"
    examples = red_snippets[:3] or neg_snippets[:3]

    return {
        "텍스트리뷰수": m,
        "직접조작언급": direct,
        "광고협찬언급": promo,
        "강한불만리뷰": severe,
        "일반불만리뷰": negative,
        "긍정리뷰": positive,
        "텍스트감점": text_penalty,
        "텍스트보너스": text_bonus,
        "텍스트신호점수": signal,
        "텍스트요약": summary,
        "대표문구": " / ".join(examples),
    }


def build_text_rankings(xlsx_path: Path, k: float, weight_cap: float, penalty_scale: float, rating_weight: float, rating_baseline: float):
    workbook = read_workbook(xlsx_path)
    places = rows_as_dicts(workbook["식당요약"])
    reviews = rows_as_dicts(workbook["리뷰전체"])
    base = compute_rankings(places, reviews, k, weight_cap, penalty_scale, rating_weight, rating_baseline)

    grouped = defaultdict(list)
    for review in reviews:
        grouped[str(review.get("식당번호", "")).strip()].append(review)

    for row in base:
        metrics = analyze_reviews(grouped.get(str(row.get("식당번호", "")).strip(), []))
        row.update(metrics)
        base_score = row.get("최종보정점수")
        if base_score is None:
            row["내용반영점수"] = None
        else:
            row["내용반영점수"] = base_score + metrics["텍스트신호점수"]

    base.sort(
        key=lambda row: (
            row["내용반영점수"] is not None,
            row["내용반영점수"] if row["내용반영점수"] is not None else -999,
            row["최종보정점수"] if row["최종보정점수"] is not None else -999,
            row["식당평균별점"] if row["식당평균별점"] is not None else -999,
            row["유효리뷰수"],
        ),
        reverse=True,
    )
    for i, row in enumerate(base, 1):
        row["순위"] = i
    return base


def write_csv(path: Path, rows: list[dict]) -> None:
    fields = [
        "순위", "장소명", "카테고리", "내용반영점수", "최종보정점수", "텍스트신호점수", "텍스트감점", "텍스트보너스",
        "텍스트요약", "대표문구", "직접조작언급", "광고협찬언급", "강한불만리뷰", "일반불만리뷰", "긍정리뷰",
        "식당평균별점", "식당총리뷰수", "캡처리뷰수", "유효리뷰수", "텍스트리뷰수", "조작의심점수", "조작감점",
        "리뷰어평균대비점수", "평균리뷰별점", "평균리뷰어평균", "원본URL",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            out = {}
            for field in fields:
                value = row.get(field, "")
                out[field] = safe_score(value) if isinstance(value, float) else value
            writer.writerow(out)


def write_html(path: Path, rows: list[dict], source: Path) -> None:
    tr = []
    for row in rows:
        score = row.get("내용반영점수")
        score_class = "positive" if (score or 0) > 0 else "negative" if score is not None else "empty"
        text_class = "risk-high" if row["텍스트감점"] >= 0.25 else "risk-mid" if row["텍스트감점"] >= 0.08 else "risk-low"
        risk_class = "risk-high" if row["조작의심점수"] >= 0.55 else "risk-mid" if row["조작의심점수"] >= 0.25 else "risk-low"
        tr.append(
            "<tr>"
            f"<td class='rank'>{row['순위']}</td>"
            f"<td class='name'><a href='{html.escape(row['원본URL'])}'>{html.escape(row['장소명'])}</a><span>{html.escape(row['카테고리'])}</span></td>"
            f"<td class='{score_class}'>{safe_score(row['내용반영점수'])}</td>"
            f"<td>{safe_score(row['최종보정점수'])}</td>"
            f"<td class='{text_class}'>{safe_score(row['텍스트감점'])}</td>"
            f"<td>{safe_score(row['텍스트보너스'])}</td>"
            f"<td class='{risk_class}'>{safe_score(row['조작의심점수'])}</td>"
            f"<td>{safe_score(row['식당평균별점'])}</td>"
            f"<td>{row['유효리뷰수']}</td>"
            f"<td>{row['텍스트리뷰수']}</td>"
            f"<td>{row['직접조작언급']}</td>"
            f"<td>{row['광고협찬언급']}</td>"
            f"<td>{row['강한불만리뷰']}</td>"
            f"<td>{row['일반불만리뷰']}</td>"
            f"<td>{row['긍정리뷰']}</td>"
            f"<td class='memo'>{html.escape(row['텍스트요약'])}</td>"
            f"<td class='memo'>{html.escape(row['대표문구'])}</td>"
            "</tr>"
        )
    content = f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>카카오 리뷰 내용 반영 순위표</title>
<style>
  body {{ margin: 0; background: #f7f8fb; color: #1f2937; font-family: "Malgun Gothic", Arial, sans-serif; }}
  main {{ width: min(1900px, calc(100vw - 32px)); margin: 28px auto 56px; }}
  h1 {{ margin: 0 0 8px; font-size: 28px; }}
  .note {{ color: #6b7280; font-size: 13px; margin: 8px 0 18px; line-height: 1.6; }}
  .formula {{ background: #eef6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 14px; line-height: 1.6; margin-bottom: 18px; }}
  table {{ width: 100%; border-collapse: collapse; background: white; border: 1px solid #d1d5db; table-layout: fixed; }}
  th, td {{ border-bottom: 1px solid #e5e7eb; padding: 8px 7px; text-align: right; vertical-align: top; font-size: 12px; }}
  th {{ position: sticky; top: 0; background: #1f4e78; color: white; z-index: 1; }}
  td.rank {{ text-align: center; font-weight: 800; width: 48px; }}
  td.name, th.name {{ text-align: left; width: 220px; }}
  td.name a {{ color: #111827; font-weight: 800; text-decoration: none; }}
  td.name span {{ display: block; margin-top: 3px; color: #6b7280; font-size: 12px; }}
  .memo {{ text-align: left; width: 260px; color: #374151; line-height: 1.45; word-break: keep-all; }}
  .positive {{ color: #047857; font-weight: 800; }}
  .negative {{ color: #b91c1c; font-weight: 800; }}
  .empty {{ color: #9ca3af; }}
  .risk-low {{ color: #047857; font-weight: 800; }}
  .risk-mid {{ color: #b45309; font-weight: 800; }}
  .risk-high {{ color: #b91c1c; font-weight: 900; }}
  @media (max-width: 900px) {{ main {{ width: 100%; margin: 0; padding: 14px; overflow-x: auto; }} table {{ min-width: 1900px; }} }}
</style>
</head>
<body>
<main>
  <h1>카카오 리뷰 내용 반영 순위표</h1>
  <div class="note">출처 파일: {html.escape(str(source))}<br>숫자 기반 점수에 실제 리뷰 텍스트 신호를 추가 반영했습니다.</div>
  <section class="formula">
    <b>내용 반영식</b><br>
    내용반영점수 = 기존 최종보정점수 - 텍스트감점 + 텍스트보너스<br>
    텍스트감점은 리뷰 알바/조작/가짜리뷰 직접 언급, 광고/협찬/체험단/리뷰이벤트 언급, 강한 불만, 일반 불만을 합산하되 상한을 둡니다.<br>
    텍스트보너스는 맛있다/친절/추천/재방문 같은 긍정 신호가 불만 신호보다 많을 때만 작게 반영합니다.
  </section>
  <table>
    <thead><tr>
      <th>순위</th><th class="name">식당</th><th>내용반영</th><th>기존최종</th><th>텍스트감점</th><th>텍스트보너스</th><th>패턴의심</th><th>평균별점</th><th>유효</th><th>텍스트</th><th>조작언급</th><th>광고협찬</th><th>강한불만</th><th>일반불만</th><th>긍정</th><th class="memo">텍스트 요약</th><th class="memo">대표 문구</th>
    </tr></thead>
    <tbody>{''.join(tr)}</tbody>
  </table>
</main>
</body>
</html>"""
    path.write_text(content, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="리뷰 내용까지 반영한 카카오 식당 순위표를 생성합니다.")
    parser.add_argument("xlsx_path", type=Path)
    parser.add_argument("--out", type=Path, default=None)
    parser.add_argument("--k", type=float, default=10.0)
    parser.add_argument("--weight-cap", type=float, default=10.0)
    parser.add_argument("--penalty-scale", type=float, default=0.5)
    parser.add_argument("--rating-weight", type=float, default=0.35)
    parser.add_argument("--rating-baseline", type=float, default=4.0)
    args = parser.parse_args()

    rows = build_text_rankings(args.xlsx_path, args.k, args.weight_cap, args.penalty_scale, args.rating_weight, args.rating_baseline)
    out = args.out or args.xlsx_path.with_name(args.xlsx_path.stem + "_내용반영순위표.html")
    out.parent.mkdir(parents=True, exist_ok=True)
    csv_path = out.with_suffix(".csv")
    write_html(out, rows, args.xlsx_path)
    write_csv(csv_path, rows)
    print(f"HTML: {out}")
    print(f"CSV:  {csv_path}")
    print(f"식당 수: {len(rows)}, 리뷰 있는 식당: {sum(1 for r in rows if r['유효리뷰수'] > 0)}, 텍스트 리뷰: {sum(r['텍스트리뷰수'] for r in rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


