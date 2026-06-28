# Kakao Review Ranker

A Python/Node.js pipeline for collecting Kakao Map restaurant reviews and generating quantitative restaurant rankings based on reviewer behavior, ratings, review volume, and suspicious rating patterns.

## What It Does

This project automates a Kakao Map restaurant review analysis workflow:

1. Search restaurants in a target region using the Kakao Local REST API.
2. Collect visible review data from Kakao Map place pages with Playwright.
3. Build a combined Excel workbook with restaurant summaries and review-level data.
4. Generate an HTML and CSV ranking table using quantitative signals.

The default pipeline does not use review text sentiment or content scoring. It focuses on numeric and behavioral signals such as:

- restaurant average rating
- reviewer average rating
- rating given to the restaurant
- reviewer review count
- reviewer follower count
- review volume
- review date concentration
- suspicious clusters of high ratings

## Project Structure

```text
xlsx_build/
  run_kakao_review_pipeline.py   # one-command pipeline runner
  kakao_restaurant_search.py     # Kakao Local API restaurant search
  kakao_place_visible_text.mjs   # Kakao Map review collector
  kakao_review_ranker.py         # quantitative ranking model
  kakao_review_text_ranker.py    # experimental text-based ranker, not used by default
  merge_retry_reviews.mjs        # helper for merging retry collections
docs/
  kakao_restaurant_analysis_timeline.md
  chat_full_work_timeline.md
```

## Requirements

- Python 3.10+
- Node.js 18+
- Kakao REST API key
- Python packages:
  - `openpyxl`
- Node packages:
  - `playwright-core`

Install Node dependencies from the project root, the folder that contains `package.json` and `xlsx_build/`:

```powershell
npm install
```

If you copied only the `xlsx_build/` folder somewhere else, install Playwright in the parent folder of `xlsx_build/`.
For example, if the script path is `C:\Users\kimwo\OneDrive\Desktop\xlsx_build\kakao_place_visible_text.mjs`, run:

```powershell
cd C:\Users\kimwo\OneDrive\Desktop
npm install playwright-core
```

The review collector launches your installed Chrome or Microsoft Edge. If neither is installed in the default location, set `CHROME_PATH` manually:

```powershell
$env:CHROME_PATH="C:\Program Files\Google\Chrome\Application\chrome.exe"
```

Install Python dependency if needed:

```powershell
pip install openpyxl
```

## Kakao REST API Key

Set your Kakao REST API key in PowerShell:

```powershell
$env:KAKAO_REST_API_KEY="YOUR_KAKAO_REST_API_KEY"
```

Check whether it is set:

```powershell
$env:KAKAO_REST_API_KEY
```

If you do not want to print the key itself, check only its length:

```powershell
$env:KAKAO_REST_API_KEY.Length
```

## Quick Start

Run the full quantitative pipeline:

```powershell
python xlsx_build\run_kakao_review_pipeline.py "인천광역시 연수구 송도과학로"
```

Edit the region inside the quotes to search another area.

```powershell
python xlsx_build\run_kakao_review_pipeline.py "YOUR_TARGET_REGION"
```

For example, replace only this part:

```text
"인천광역시 연수구 송도과학로"
```

with something like:

```text
"서울특별시 마포구 연남동"
"부산광역시 해운대구 우동"
"대전광역시 유성구 궁동"
```

Or run it interactively:

```powershell
python xlsx_build\run_kakao_review_pipeline.py
```

Test with only the first 5 restaurants:

```powershell
python xlsx_build\run_kakao_review_pipeline.py "인천광역시 연수구 송도과학로" --limit 5
```

Results are written inside the `xlsx_build` folder:

```text
xlsx_build/outputs/review_pipeline/<region_timestamp>/
```

The folder name is generated automatically from the search region and run time.
Spaces and symbols in the region are converted to underscores.

Example:

```text
xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/
```

Each run folder contains:

- `restaurants/kakao_restaurants_<region>_<timestamp>.xlsx`: restaurant search results
- `restaurants/kakao_restaurants_<region>_<timestamp>.csv`: restaurant search results as CSV
- `restaurants/kakao_restaurants_<region>_<timestamp>.json`: restaurant search results as JSON
- `raw_place_text/*.txt`: raw visible text collected from Kakao Map place pages
- `raw_place_text/*.json`: parsed collection metadata for each place
- `<region>_통합리뷰.xlsx`: combined review workbook
- `<region>_정량순위표.html`: quantitative ranking report
- `<region>_정량순위표.csv`: quantitative ranking table as CSV

For the example region above, the main output files look like this:

```text
xlsx_build/xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/인천광역시_연수구_송도과학로_통합리뷰.xlsx
xlsx_build/xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/인천광역시_연수구_송도과학로_정량순위표.html
xlsx_build/xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/인천광역시_연수구_송도과학로_정량순위표.csv
```

## Manual Workflow

You can also run each step separately.

Search restaurants:

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도과학로"
```

Collect reviews into one workbook:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\restaurants.xlsx" --combined-xlsx "outputs\kakao_place_text\combined_reviews.xlsx"
```

Generate a quantitative ranking:

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\combined_reviews.xlsx" --out "outputs\kakao_place_text\ranking.html"
```

## GitHub Usage

Recommended repository name:

```text
kakao-review-ranker
```

Recommended short description:

```text
Collect and rank Kakao Map restaurant reviews using quantitative rating and reviewer signals.
```

For normal use, keep generated outputs out of Git and commit only code, documentation, and small sample files. If you want to share analysis results, export a selected HTML/CSV report separately or place a small anonymized sample under a dedicated `examples/` folder.

## Notes

Kakao Map page structure can change over time, so review collection may need maintenance if the visible text format changes. Always check the generated workbook before relying on a ranking result.

## Troubleshooting

### Playwright not found

If review collection fails with this message:

```text
Error: Playwright를 찾지 못했습니다. npm install playwright-core 또는 npm install playwright가 필요합니다.
```

install the Node dependency in the folder above `xlsx_build/`:

```powershell
cd C:\path\to\folder_that_contains_xlsx_build
npm install playwright-core
```

For the full repository, this usually means:

```powershell
cd C:\Users\kimwo\project_review
npm install
```

For a copied Desktop-only setup like `C:\Users\kimwo\OneDrive\Desktop\xlsx_build`, this means:

```powershell
cd C:\Users\kimwo\OneDrive\Desktop
npm install playwright-core
```

Then run the pipeline again.