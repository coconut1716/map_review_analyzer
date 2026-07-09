# Kakao Review Ranker

Kakao Map 음식점 리뷰를 수집하고, 리뷰어 행동·별점·리뷰 수·날짜 집중도 같은 정량 신호로 음식점 순위를 만드는 Python/Node.js 파이프라인입니다.

> **중요:** 이 저장소는 GitHub에 올리기 위해 생성 결과물과 로컬 의존성(`node_modules/`, `outputs/`, `*.xlsx`, `*.csv`, `*.json`, `*.html` 등)을 `.gitignore`로 제외합니다. 따라서 Git에서 보이는 파일만 내려받은 뒤에는 아래의 **설치해야 하는 것**을 먼저 준비해야 실행할 수 있습니다.

## 현재 Git에 포함된 파일

```text
README.md
package.json
package-lock.json
xlsx_build/
  run_kakao_review_pipeline.py
  kakao_restaurant_search.py
  kakao_place_visible_text.mjs
  kakao_review_ranker.py
  kakao_review_text_ranker.py
  merge_retry_reviews.mjs
  package.json
  readme.md
docs/
  kakao_restaurant_analysis_timeline.md
```

## Git에 포함되지 않는 파일

아래 파일과 폴더는 용량, 개인정보, 실행 환경 차이 때문에 Git에 올리지 않습니다.

```text
node_modules/                 # npm install로 다시 설치
xlsx_build/node_modules/      # xlsx_build 단독 실행용 로컬 의존성
outputs/                      # 실행 결과
xlsx_build/outputs/           # 파이프라인 실행 결과
*.xlsx, *.csv, *.json, *.html # 수집/분석 결과 파일
.env, *.env                   # API 키 같은 비밀값
```

즉, 기존 로컬 PC에 있던 결과 엑셀·CSV·HTML 파일은 GitHub에서 다시 받는 파일이 아니라, 실행 후 새로 생성되는 파일입니다.

## 무엇을 하는 프로젝트인가요?

기본 파이프라인은 다음 작업을 자동으로 이어서 실행합니다.

1. Kakao Local REST API로 대상 지역의 음식점을 검색합니다.
2. Playwright로 Kakao Map 장소 페이지에서 보이는 리뷰 정보를 수집합니다.
3. 음식점 요약과 리뷰 단위 데이터를 하나의 Excel 워크북으로 만듭니다.
4. 정량 지표 기반 HTML/CSV 순위표를 생성합니다.

기본 분석은 리뷰 본문 감성 분석을 사용하지 않습니다. 대신 다음 수치를 활용합니다.

- 식당 평균 별점
- 리뷰어 평균 별점
- 해당 식당에 준 별점
- 리뷰어 리뷰 수
- 리뷰어 팔로워 수
- 식당 리뷰 수
- 리뷰 날짜 집중도
- 높은 별점의 의심스러운 집중 패턴

## 설치해야 하는 것

### 1. 시스템 프로그램

- Python 3.10 이상
- Node.js 18 이상
- Google Chrome 또는 Microsoft Edge
- Kakao REST API 키

### 2. Python 패키지

음식점 검색 결과를 XLSX로 저장하려면 `openpyxl`이 필요합니다.

```powershell
pip install openpyxl
```

### 3. Node 패키지

저장소 루트에서 `package-lock.json` 기준으로 설치합니다.

```powershell
npm install
```

현재 루트 `package.json`에는 Kakao Map 페이지 수집에 필요한 Playwright 런타임인 `playwright-core`가 들어 있습니다.

```text
playwright-core
```

### 4. XLSX 읽기/쓰기 도구 관련 주의

`xlsx_build/kakao_place_visible_text.mjs`와 `xlsx_build/merge_retry_reviews.mjs`는 XLSX 파일을 읽고 쓰기 위해 `@oai/artifact-tool` 모듈을 import합니다.

이 모듈은 현재 Git에 포함되어 있지 않은 `node_modules/`에 있던 실행 의존성입니다. GitHub에서 새로 clone한 환경에서 아래 오류가 나면, 기존 실행 환경의 `node_modules`를 복원하거나 해당 모듈을 설치할 수 있는 사내/로컬 npm 설정을 먼저 준비해야 합니다.

```text
Error [ERR_MODULE_NOT_FOUND]: Cannot find package '@oai/artifact-tool'
```

정리하면 GitHub clone만으로 준비해야 하는 실행 의존성은 다음과 같습니다.

```text
Python: openpyxl
Node: playwright-core, @oai/artifact-tool
Browser: Chrome 또는 Microsoft Edge
Secret: KAKAO_REST_API_KEY
```

## Kakao REST API 키 설정

PowerShell에서 현재 세션에 Kakao REST API 키를 설정합니다.

```powershell
$env:KAKAO_REST_API_KEY="YOUR_KAKAO_REST_API_KEY"
```

설정 여부를 확인합니다.

```powershell
$env:KAKAO_REST_API_KEY
```

키를 화면에 그대로 출력하고 싶지 않으면 길이만 확인합니다.

```powershell
$env:KAKAO_REST_API_KEY.Length
```

## 브라우저 경로 설정

리뷰 수집기는 설치된 Chrome 또는 Microsoft Edge를 실행합니다. 기본 위치에서 찾지 못하면 `CHROME_PATH`를 직접 지정합니다.

```powershell
$env:CHROME_PATH="C:\Program Files\Google\Chrome\Application\chrome.exe"
```

## 빠른 실행

저장소 루트에서 전체 정량 파이프라인을 실행합니다. 기본 식당 수집 방식은 기준 장소 좌표 주변을 격자로 훑는 `grid` 모드입니다. `--outer-radius`는 기준 장소에서 최종 유지할 식당 반경(m)입니다. 기본값은 300m이며, 아래 예시는 500m까지 수집합니다.

```powershell
python xlsx_build\run_kakao_review_pipeline.py "인천광역시 연수구 송도과학로" --outer-radius 500
```

격자 검색 기본값은 `--grid-step 100`, `--cell-radius 120`입니다. 필요하면 직접 지정할 수 있습니다.

```powershell
python xlsx_build\run_kakao_review_pipeline.py "성북동 주민센터" --outer-radius 500 --grid-step 100 --cell-radius 120 --minimal
```

다른 지역을 검색하려면 따옴표 안의 지역명만 바꿉니다.

```powershell
python xlsx_build\run_kakao_review_pipeline.py "서울특별시 마포구 연남동" --outer-radius 500
python xlsx_build\run_kakao_review_pipeline.py "부산광역시 해운대구 우동" --outer-radius 500
python xlsx_build\run_kakao_review_pipeline.py "대전광역시 유성구 궁동" --outer-radius 500
```

지역명을 실행 중에 입력하려면 인자 없이 실행합니다.

```powershell
python xlsx_build\run_kakao_review_pipeline.py
```

테스트로 앞 5개 식당만 수집하려면 `--limit`을 사용합니다.

```powershell
python xlsx_build\run_kakao_review_pipeline.py "인천광역시 연수구 송도과학로" --outer-radius 500 --limit 5
```

## 결과 파일 위치

전체 파이프라인 결과는 `xlsx_build/outputs/review_pipeline/` 아래에 생성됩니다.

```text
xlsx_build/outputs/review_pipeline/<지역명_실행시간>/
```

예:

```text
xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/
```

각 실행 폴더에는 보통 다음 파일이 생깁니다.

```text
restaurants/kakao_restaurants_<지역명>_<시간>.xlsx
restaurants/kakao_restaurants_<지역명>_<시간>.csv
restaurants/kakao_restaurants_<지역명>_<시간>.json
raw_place_text/*.txt
raw_place_text/*.json
<지역명>_통합리뷰.xlsx
<지역명>_정량순위표.html
<지역명>_정량순위표.csv
```

예:

```text
xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/인천광역시_연수구_송도과학로_통합리뷰.xlsx
xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/인천광역시_연수구_송도과학로_정량순위표.html
xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260627_143012/인천광역시_연수구_송도과학로_정량순위표.csv
```

## 수동 실행

각 단계를 따로 실행할 수도 있습니다.

### 1. 음식점 검색

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도과학로"
```

기본 결과 위치:

```text
outputs/kakao_restaurants/
```

### 2. 리뷰 수집

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\restaurants.xlsx" --combined-xlsx "outputs\kakao_place_text\combined_reviews.xlsx"
```

### 3. 정량 순위표 생성

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\combined_reviews.xlsx" --out "outputs\kakao_place_text\ranking.html"
```

## 문제 해결

### Playwright를 찾지 못하는 경우

오류 예:

```text
Error: Playwright를 찾지 못했습니다. npm install playwright-core 또는 npm install playwright가 필요합니다.
```

저장소 루트에서 Node 의존성을 다시 설치합니다.

```powershell
npm install
```

### @oai/artifact-tool을 찾지 못하는 경우

오류 예:

```text
Error [ERR_MODULE_NOT_FOUND]: Cannot find package '@oai/artifact-tool'
```

현재 Git에 보이는 파일만으로는 이 모듈이 포함되지 않습니다. 기존 로컬 실행 환경의 `node_modules`를 복원하거나, 해당 패키지를 설치할 수 있는 npm 레지스트리/인증 설정을 준비해야 합니다.

### XLSX 파일이 생성되지 않는 경우

`openpyxl`이 없으면 음식점 검색 단계에서 XLSX 저장이 생략될 수 있습니다.

```powershell
pip install openpyxl
```

### 실행 결과 파일이 Git에 보이지 않는 경우

정상입니다. `.gitignore`가 `outputs/`, `*.xlsx`, `*.csv`, `*.json`, `*.html` 같은 생성 파일을 제외합니다. 결과를 공유하려면 필요한 HTML/CSV 파일만 별도로 전달하거나, 개인정보를 제거한 작은 예시 파일을 별도 폴더에 추가하는 방식을 권장합니다.

## 참고

Kakao Map 페이지 구조는 시간이 지나면 바뀔 수 있습니다. 리뷰 수집 결과가 비어 있거나 형식이 달라지면 `kakao_place_visible_text.mjs`의 파싱 로직을 점검해야 합니다.
