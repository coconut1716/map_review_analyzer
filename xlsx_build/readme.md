# Kakao Review Ranker Tools

이 폴더(`xlsx_build`)만 있어도 카카오맵 음식점 수집부터 정량 리뷰 분석까지 실행할 수 있게 구성되어 있다.

## 폴더 구성

```text
xlsx_build/
  run_kakao_review_pipeline.py   # 한 번에 실행하는 메인 파일
  kakao_restaurant_search.py     # Kakao Local API 음식점 검색
  kakao_place_visible_text.mjs   # Kakao Map 리뷰 수집
  kakao_review_ranker.py         # 정량 순위표 생성
  kakao_review_text_ranker.py    # 텍스트 분석 실험용, 기본 파이프라인에서는 사용 안 함
  merge_retry_reviews.mjs        # 재수집 결과 병합 도우미
  node_modules/                  # Node 실행에 필요한 로컬 의존성
  package.json                   # xlsx_build 단독 설치용
```

## 필요한 것

- Python 3.10+
- Node.js 18+
- Kakao REST API key
- Chrome 또는 Microsoft Edge
- Python 패키지 `openpyxl`

Python 패키지가 없으면 설치한다.

```powershell
pip install openpyxl
```

## Node 의존성

이 폴더를 통째로 복사할 때는 `node_modules/`도 같이 복사해야 완전 단독 실행된다.

특히 `@oai/artifact-tool`은 XLSX 파일을 읽고 쓰는 데 필요하므로, `node_modules/`를 삭제하면 단독 실행이 깨질 수 있다.

Playwright만 빠져서 오류가 나는 경우에는 `xlsx_build` 폴더 안에서 설치한다.

```powershell
cd C:\path\to\xlsx_build
npm install
```

오류 메시지 예:

```text
Error: Playwright를 찾지 못했습니다. npm install playwright-core 또는 npm install playwright가 필요합니다.
```

이 경우에는 위의 `npm install`을 `xlsx_build` 안에서 실행한다. 그래도 `@oai/artifact-tool` 관련 오류가 나면, 원본 `xlsx_build/node_modules` 폴더를 다시 함께 복사해야 한다.

## REST API 키 설정

PowerShell에서 현재 세션에 Kakao REST API 키를 설정한다.

```powershell
$env:KAKAO_REST_API_KEY="YOUR_KAKAO_REST_API_KEY"
```

설정 확인:

```powershell
$env:KAKAO_REST_API_KEY
```

키를 화면에 그대로 보이기 싫으면 길이만 확인한다.

```powershell
$env:KAKAO_REST_API_KEY.Length
```

## 실행 방법

`xlsx_build` 폴더 안에서 실행한다. 기본 식당 수집 방식은 기준 장소 좌표 주변을 격자로 훑는 `grid` 모드다. `--outer-radius`는 기준 장소에서 최종 유지할 식당 반경(m)이며 기본값은 300m다. 아래 예시는 500m까지 수집한다.

```powershell
cd C:\path\to\xlsx_build
python run_kakao_review_pipeline.py "인천광역시 연수구 송도과학로" --outer-radius 500
```

격자 검색 기본값은 `--grid-step 100`, `--cell-radius 120`이다. 필요하면 직접 지정할 수 있다.

```powershell
python run_kakao_review_pipeline.py "성북동 주민센터" --outer-radius 500 --grid-step 100 --cell-radius 120 --minimal
```

다른 지역을 검색하려면 따옴표 안의 지역명만 바꾼다.

```powershell
python run_kakao_review_pipeline.py "서울특별시 마포구 연남동" --outer-radius 500
```

지역명을 실행 중에 입력하고 싶으면:

```powershell
python run_kakao_review_pipeline.py
```

테스트로 앞 5개 식당만 수집하려면:

```powershell
python run_kakao_review_pipeline.py "인천광역시 연수구 송도과학로" --outer-radius 500 --limit 5
```

## 결과 파일 위치

결과는 `xlsx_build` 폴더 안의 `outputs/review_pipeline/` 아래에 생성된다.

```text
xlsx_build/outputs/review_pipeline/<지역명_실행시간>/
```

예:

```text
xlsx_build/outputs/review_pipeline/인천광역시_연수구_송도과학로_20260628_131500/
```

주요 결과 파일:

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
outputs/review_pipeline/인천광역시_연수구_송도과학로_20260628_131500/인천광역시_연수구_송도과학로_통합리뷰.xlsx
outputs/review_pipeline/인천광역시_연수구_송도과학로_20260628_131500/인천광역시_연수구_송도과학로_정량순위표.html
outputs/review_pipeline/인천광역시_연수구_송도과학로_20260628_131500/인천광역시_연수구_송도과학로_정량순위표.csv
```

## 현재 기본 분석 방식

기본 파이프라인은 `kakao_review_ranker.py`만 사용한다.
리뷰 내용 문장 분석은 하지 않고, 다음 정량 수치를 활용한다.

- 식당 평균 별점
- 리뷰어 평균 별점
- 이 식당에 준 별점
- 리뷰어 총 리뷰 수
- 팔로워 수
- 리뷰 날짜 집중도
- 별 5개 리뷰 집중/버스트 패턴

`kakao_review_text_ranker.py`는 실험용이며 기본 실행에는 연결되어 있지 않다.