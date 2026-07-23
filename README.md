# Kakao Review Ranker

Kakao Map의 음식점 리뷰를 수집하고, 리뷰어의 평소 평가 성향과 활동량을 반영해 음식점 순위를 만드는 프로젝트입니다. 단순 평균 별점은 리뷰어마다 점수를 주는 기준이 다르고 리뷰 이벤트나 짧은 기간의 리뷰 집중에 영향을 받을 수 있다는 문제에서 시작했습니다.

## 실행 환경과 의존성

- [Python 3.10+](https://www.python.org/downloads/) — 검색, 데이터 정리, 순위 계산
- [Node.js 18+](https://nodejs.org/) — Kakao Map 리뷰 수집기 실행
- [Playwright](https://playwright.dev/docs/intro) — 브라우저 자동화
- [openpyxl](https://openpyxl.readthedocs.io/) — 검색 결과를 XLSX로 저장
- Chrome 또는 Microsoft Edge
- [Kakao Local REST API 키](https://developers.kakao.com/docs/latest/ko/local/common)

<details>
<summary><strong>설치 방법과 모듈별 역할 보기</strong></summary>

```bash
python -m pip install openpyxl
npm ci
```

`xlsx_build/kakao_place_visible_text.mjs`는 XLSX 처리를 위해 `@oai/artifact-tool`도 사용합니다. 이 모듈은 공개 저장소의 `package.json`에 포함되어 있지 않으므로, 실행 환경에서 별도로 제공되어야 합니다. 모듈을 찾을 수 없다는 오류가 발생하면 해당 패키지를 사용할 수 있는 npm 환경이나 기존 실행 환경의 의존성이 필요합니다.

| 파일 | 역할 |
|---|---|
| `kakao_restaurant_search.py` | Kakao Local API로 음식점 검색 |
| `kakao_place_visible_text.mjs` | 음식점 페이지의 리뷰 수집 |
| `kakao_review_ranker.py` | 정량 점수 계산 및 HTML/CSV 생성 |
| `run_kakao_review_pipeline.py` | 위 과정을 한 번에 실행 |

</details>

## 작동 원리

1. 입력한 지역을 격자로 나누어 주변 음식점을 검색합니다.
2. 각 음식점의 별점과 공개된 리뷰어 정보를 수집합니다.
3. 리뷰마다 `이 식당에 준 별점 - 리뷰어의 평균 별점`을 계산합니다.
4. 리뷰 수와 팔로워 수로 구한 리뷰어 신뢰도 가중치를 이 차이에 적용합니다.
5. 값들을 합산한 뒤 전체 가중치의 합으로 나누어 리뷰어 성향 대비 점수를 구합니다.
6. 유효 리뷰가 적은 식당은 `n / (n + k)`로 점수를 보수적으로 낮춥니다.
7. 리뷰 날짜 집중, 반복 주기, 저활동 계정의 고평점 비율 등으로 조작 의심도를 계산해 감점합니다.
8. 마지막으로 조작 의심도가 낮을 때만 식당 평균 별점을 일부 반영해 최종 순위를 만듭니다.

## 실행 방법

저장소를 내려받은 뒤 루트 디렉터리에서 의존성을 설치하고 API 키를 환경 변수로 설정합니다.

```bash
# macOS / Linux
export KAKAO_REST_API_KEY="YOUR_KAKAO_REST_API_KEY"

# Windows PowerShell
$env:KAKAO_REST_API_KEY="YOUR_KAKAO_REST_API_KEY"
```

전체 파이프라인은 다음 명령으로 실행합니다. `<검색 지역>`에는 주소, 동네 이름 또는 기준 장소를 입력할 수 있습니다.

```bash
python xlsx_build/run_kakao_review_pipeline.py "<검색 지역>" --outer-radius 500
```

먼저 일부 식당만 시험하려면 `--limit`을 사용합니다.

```bash
python xlsx_build/run_kakao_review_pipeline.py "<검색 지역>" --outer-radius 500 --limit 5
```

결과는 `xlsx_build/outputs/review_pipeline/<지역명_실행시간>/`에 XLSX, CSV, HTML 형식으로 생성됩니다. 전체 옵션은 `python xlsx_build/run_kakao_review_pipeline.py --help`에서 확인할 수 있습니다.

## 현재 결론과 다음 목표

이 프로젝트는 평균 별점을 정답으로 간주하기보다, 리뷰어별 평가 성향과 활동 이력, 리뷰 발생 패턴을 함께 살펴 더 설명 가능한 음식점 순위를 만드는 실험입니다. 앞으로는 적절한 대기 시간과 `stall` 횟수를 찾아 수집 속도를 높이고, 실제 데이터 분포를 바탕으로 리뷰 조작 의심도 계산을 계속 개선할 계획입니다.

리뷰가 10개 이하인 식당은 표본이 작으므로 네이버 지도나 Google Maps와 교차 검증하는 방안도 고려하고 있습니다. 다만 두 서비스에서는 리뷰어의 평균 별점을 동일한 방식으로 얻기 어렵고, 리뷰 수가 많더라도 리뷰 이벤트의 영향을 받을 수 있습니다. 블로그 리뷰 수 역시 레뷰·강남맛집·서울오빠 등 체험단 및 리뷰 마케팅 플랫폼의 영향을 구분하기 어려워 독립적인 신뢰 지표로 사용하기에는 한계가 있습니다. 따라서 다른 지도 서비스의 평균 별점이나 리뷰 수는 직접적인 정답이 아니라 보조 신호로 검토할 예정입니다.

> Kakao Map의 페이지 구조가 바뀌면 리뷰 수집 및 파싱 로직도 수정해야 할 수 있습니다.
