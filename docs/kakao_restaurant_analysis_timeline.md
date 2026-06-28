# 카카오맵 식당 리뷰 분석 프로젝트 작업 기록

이 문서는 이 채팅에서 진행한 **카카오맵 식당 수집, 리뷰 수집, 리뷰 신뢰도 보정, 조작 의심 보정, 순위표 생성** 작업을 시간순으로 정리한 것이다.

작성 기준 파일:

- `xlsx_build/kakao_restaurant_search.py`
- `xlsx_build/kakao_place_visible_text.mjs`
- `xlsx_build/kakao_review_ranker.py`
- `outputs/kakao_restaurants/*.xlsx`
- `outputs/kakao_place_text/*_통합리뷰.xlsx`
- `outputs/kakao_place_text/*_순위표_*.html`

---

## 1. 목표

처음 목표는 특정 지역의 음식점을 카카오맵에서 모으고, 각 식당의 리뷰를 수집한 뒤, 단순 평균 별점이 아니라 다음 요소를 반영해서 “좋은 식당일 가능성”을 점수화하는 것이었다.

- 식당 평균 별점
- 리뷰어가 평소 주는 평균 별점
- 해당 식당에 준 별점
- 리뷰어의 총 리뷰 수
- 리뷰어의 팔로워 수
- 리뷰 날짜
- 리뷰가 특정 날짜나 특정 기간에 몰렸는지 여부
- 별 5개 리뷰가 비정상적으로 몰렸는지 여부

핵심 아이디어는 다음이다.

> 그냥 별점이 높은 식당보다, 까다로운 리뷰어들이 평소보다 높게 평가한 식당을 더 좋게 보고, 동시에 별 5개 리뷰가 특정 시기에 몰리는 경우에는 조작 의심 감점을 준다.

---

## 2. 전체 파이프라인

```text
지역명 입력
  ↓
Kakao Local REST API로 음식점 목록 수집
  ↓
place.map.kakao.com/{placeId}#review 링크 생성
  ↓
Playwright로 각 리뷰 페이지 접속
  ↓
화면에 보이는 텍스트 저장 및 파싱
  ↓
식당요약 / 리뷰전체 통합 XLSX 생성
  ↓
리뷰어 신뢰도, 리뷰어 평균 대비 점수, 조작 의심 점수 계산
  ↓
CSV + HTML 순위표 생성
```

---

## 3. Timeline

## 3.1 카카오맵 데이터 접근 방식 검토

처음에는 카카오맵 웹사이트에서 직접 검색 결과와 리뷰를 긁는 방식을 고민했다.

하지만 카카오맵은 내부 검색을 해도 주소창 URL이 잘 바뀌지 않고, 검색 결과가 프론트엔드 내부 상태로 관리된다. 그래서 단순 URL 패턴만으로는 특정 지역 음식점 목록을 안정적으로 얻기 어렵다.

이때 두 갈래를 검토했다.

| 방식 | 장점 | 단점 |
|---|---|---|
| 웹 UI 직접 스크래핑 | API 키가 필요 없음 | UI 변경에 취약, 검색 결과 스크롤/클릭이 복잡 |
| Kakao Local REST API | 지역 음식점 목록과 place URL을 안정적으로 얻음 | REST API 키 필요 |

최종적으로 음식점 목록은 **Kakao Local REST API**로 받고, 리뷰는 카카오 플레이스 페이지를 브라우저로 열어 visible text를 가져오는 혼합 방식을 선택했다.

---

## 3.2 Kakao Developers REST API 키 설정

카카오 개발자 콘솔에서 앱을 생성하고 REST API 키를 발급했다.

중요했던 점:

- 앱 생성 후 `OPEN_MAP_AND_LOCAL` 서비스가 비활성화되어 있으면 API가 403을 반환한다.
- 실제 에러는 다음 형태였다.

```text
HTTP 403
NotAuthorizedError
App disabled OPEN_MAP_AND_LOCAL service.
```

해결:

- Kakao Developers 앱 설정에서 지도/로컬 관련 서비스를 활성화했다.
- PowerShell에서는 환경변수를 다음처럼 설정한다.

```powershell
$env:KAKAO_REST_API_KEY="REST_API_KEY_값"
```

주의:

- `cmd.exe`에서는 `$env:` 문법이 동작하지 않는다.
- PowerShell에서 환경변수 설정은 성공해도 별도 성공 메시지가 뜨지 않는다.
- 에러가 안 뜨면 정상으로 보면 된다.

---

## 3.3 음식점 목록 수집기 추가

파일:

```text
xlsx_build/kakao_restaurant_search.py
```

역할:

- 특정 지역명을 입력받는다.
- Kakao Local REST API로 음식점 목록을 검색한다.
- 각 음식점의 카카오 플레이스 URL과 리뷰 URL을 만든다.
- JSON, CSV, XLSX로 저장한다.

대표 실행 명령:

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도동"
```

또는:

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도과학로"
```

주요 옵션:

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도동" --mode both --radius 2000 --max-pages 45 --size 15
```

### 주요 함수

### `request_json()`

Kakao API에 요청을 보내고 JSON을 읽는다.

- Authorization 헤더에 `KakaoAK {REST_API_KEY}`를 넣는다.
- HTTP 에러가 나면 상태 코드와 응답 본문을 보여주도록 수정했다.
- 403 원인 확인에 도움이 되었다.

### `search_keyword()`

키워드 검색 API를 사용한다.

예를 들어:

```text
인천 연수구 송도동 음식점
```

같은 키워드를 검색한다.

### `resolve_center()`

지역명을 주소 검색 API로 넣어 중심 좌표를 얻는다.

### `search_food_category()`

카테고리 검색 API를 사용한다.

- 음식점 카테고리 코드: `FD6`
- 중심 좌표 기준 반경 검색

### `normalize()`

Kakao API 응답을 우리가 쓰기 좋은 형태로 바꾼다.

특히 이 부분이 중요하다.

```python
place_url = doc.get("place_url", "")
review_url = f"{place_url.split('#', 1)[0]}#review" if place_url else ""
```

카카오맵 장소 링크를 리뷰 탭 링크로 바꾼다.

예:

```text
https://place.map.kakao.com/450422239
```

을:

```text
https://place.map.kakao.com/450422239#review
```

로 만든다.

### 출력 컬럼

대표 컬럼은 다음과 같다.

| 컬럼 | 의미 |
|---|---|
| `id` | 카카오 장소 ID |
| `place_name` | 장소명 |
| `category_name` | 카테고리 |
| `phone` | 전화번호 |
| `address_name` | 지번 주소 |
| `road_address_name` | 도로명 주소 |
| `place_url` | 카카오 장소 URL |
| `review_url` | 리뷰 탭 URL |
| `region_query` | 검색한 지역명 |
| `source` | keyword/category 검색 출처 |

### 생성된 결과 예시

```text
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx
```

---

## 3.4 카카오 플레이스 리뷰 텍스트 수집기 추가

파일:

```text
xlsx_build/kakao_place_visible_text.mjs
```

초기 목표는 다음 페이지에서 보이는 리뷰 텍스트를 가져오는 것이었다.

```text
https://place.map.kakao.com/450422239#review
```

처음에는 “브라우저에서 Ctrl+A, Ctrl+C로 복사한 것처럼 보이는 텍스트를 가져오면 되지 않을까?”라는 방향이었다.

최종 구현은 Playwright로 페이지에 접속하고, 스크롤한 뒤, 페이지의 visible text를 저장하는 방식이다.

### 대표 실행 명령

단일 URL:

```powershell
node xlsx_build\kakao_place_visible_text.mjs "https://place.map.kakao.com/450422239#review" --scrolls 10
```

음식점 목록 XLSX의 모든 `review_url`에 대해 실행:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --scrolls 3 --fast-threshold 40 --max-scrolls 200 --target-coverage 0.95 --stall-limit 8 --stall-wait-ms 5000 --wait-ms 2500 --scroll-delay-ms 1200 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx"
```

처음 10개만 테스트:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --limit 10 --scrolls 10 --wait-ms 1500 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_10개.xlsx"
```

### 주요 옵션

| 옵션 | 의미 |
|---|---|
| `--xlsx-file` | XLSX에서 URL 열을 읽는다 |
| `--url-column` | URL 열 이름. 기본값 `review_url` |
| `--limit` | 앞 N개만 실행 |
| `--scrolls` | 리뷰 페이지 스크롤 횟수 |
| `--wait-ms` | 페이지 진입 후 대기 시간 |
| `--combined-xlsx` | 모든 식당 리뷰를 하나의 XLSX로 저장 |
| `--txt-file` | 이미 저장된 TXT를 다시 파싱 |
| `--screenshot` | 페이지 스크린샷 저장 |
| `--preview` | XLSX 미리보기 PNG 저장 |

---

## 3.5 리뷰 텍스트 파싱 로직 추가

파일:

```text
xlsx_build/kakao_place_visible_text.mjs
```

### `parsePlaceText()`

카카오 플레이스에서 저장한 visible text를 분석해 식당 정보와 리뷰어 정보를 뽑는다.

파싱한 식당 정보:

| 항목 | 설명 |
|---|---|
| 장소명 | 식당 이름 |
| 카테고리 | 음식점 종류 |
| 식당 평균 별점 | 카카오에 표시된 평균 별점 |
| 식당 총 리뷰수 | 카카오 후기 수 |
| 블로그 수 | 블로그 리뷰 수 |
| 원본 URL | 카카오 플레이스 URL |

파싱한 리뷰어 정보:

| 항목 | 설명 |
|---|---|
| 익명 리뷰어 | `익명1`, `익명2`처럼 저장 |
| 리뷰 날짜 | 리뷰 작성일 |
| 리뷰어 평균 별점 | 그 리뷰어가 평소 주는 평균 별점 |
| 이 식당에 준 별점 | 해당 식당에 준 별점 |
| 리뷰어 총 리뷰수 | 그 리뷰어가 작성한 후기 수 |
| 팔로워 수 | 리뷰어 팔로워 수 |
| 리뷰 내용 | 텍스트 리뷰 |

익명화는 다음처럼 처리했다.

```javascript
anonymousReviewer: `익명${reviews.length + 1}`
```

즉, 원래 카카오 닉네임을 저장하지 않고 `익명1`, `익명2`처럼만 남긴다.

---

## 3.6 개별 XLSX에서 통합 XLSX로 변경

처음에는 URL 하나를 실행할 때마다 식당별 XLSX가 만들어졌다.

하지만 사용자가 원한 형태는:

> 식당마다 다른 XLSX가 아니라, 모든 식당 리뷰를 하나의 XLSX에 모으는 것

그래서 `--combined-xlsx` 옵션을 추가했다.

### 통합 XLSX 구조

생성 파일 예:

```text
outputs/kakao_place_text/송도과학로_통합리뷰.xlsx
```

시트는 2개다.

### `식당요약`

| 컬럼 | 의미 |
|---|---|
| 식당번호 | 내부 순번 |
| 장소명 | 식당명 |
| 카테고리 | 카카오 카테고리 |
| 식당 평균 별점 | 카카오 표시 평균 별점 |
| 식당 총 리뷰수 | 카카오 표시 후기 수 |
| 캡처된 리뷰수 | 실제 파싱된 리뷰 수 |
| 블로그 수 | 블로그 리뷰 수 |
| 원본 URL | 카카오 플레이스 URL |
| 원본 TXT | 저장된 텍스트 파일 경로 |

### `리뷰전체`

| 컬럼 | 의미 |
|---|---|
| 식당번호 | `식당요약`과 연결되는 번호 |
| 장소명 | 식당명 |
| 식당 평균 별점 | 식당 평균 별점 |
| 식당 총 리뷰수 | 식당 총 리뷰수 |
| 익명 리뷰어 | 익명화된 리뷰어 |
| 리뷰 날짜 | 리뷰 날짜 |
| 리뷰어 평균 별점 | 리뷰어의 평소 평균 별점 |
| 이 식당에 준 별점 | 해당 식당에 준 별점 |
| 리뷰어 총 리뷰수 | 리뷰어 총 리뷰 수 |
| 팔로워 수 | 리뷰어 팔로워 수 |
| 리뷰 내용 | 리뷰 텍스트 |
| 원본 URL | 카카오 플레이스 URL |

---

## 3.7 순위표 생성기 추가

파일:

```text
xlsx_build/kakao_review_ranker.py
```

역할:

- 통합 리뷰 XLSX를 읽는다.
- 식당별 리뷰를 묶는다.
- 리뷰어 평균 대비 점수를 계산한다.
- 리뷰어 신뢰도 가중치를 적용한다.
- 리뷰 수가 적을 때 과대평가되지 않도록 신뢰도 보정을 한다.
- 조작 의심 점수를 계산해 감점한다.
- HTML 순위표와 CSV를 생성한다.

대표 실행 명령:

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_별5집중보정.html"
```

---

---

## 3.8 조작 감점 계수 상향

순위표를 확인한 뒤, 조작 의심 점수의 감점 폭이 너무 후하다는 판단이 있었다.

기존 기본값은 다음이었다.

$$
\lambda=0.35
$$

즉 조작의심점수 $G_j$가 1에 가까워도 최대 0.35점만 감점했다.

하지만 리뷰 조작 의심이 강한 식당은 실제 추천 순위에서 더 강하게 내려가야 한다고 보고, 기본 감점 계수를 다음처럼 올렸다.

$$
\lambda=0.5
$$

그래서 조작 감점은 현재 다음과 같다.

$$
H_j=0.5G_j
$$

최종 점수는 다음이다.

$$
F_j=B_j-0.5G_j
$$

코드 변경 위치:

```python
parser.add_argument("--penalty-scale", type=float, default=0.5, help="조작 의심 점수 감점 계수")
```

이 변경으로 예를 들어 조작의심점수 $G_j=0.5713$인 식당은:

$$
0.5713\times0.5=0.28565
$$

만큼 감점된다. 기존 $0.35$ 기준의 약 $0.19996$ 감점보다 더 엄격하다.

---

## 3.9 팔로워 가중치를 리뷰 수 가중치와 동일하게 변경

기존 팔로워 가중치는 리뷰 수 가중치보다 약하게 들어갔다.

$$
\psi(f_i)=1+0.25\ln(1+f_i)
$$

하지만 팔로워 수 역시 리뷰어 영향력의 신뢰 신호로 볼 수 있으므로, 팔로워 가중치도 리뷰어 총 리뷰 수와 같은 누진 함수를 쓰도록 바꿨다.

현재는 다음과 같다.

$$
\psi(f_i)=\phi(f_i)
$$

따라서 최종 리뷰어 가중치는:

$$
w_i=\min\left(10,\phi(n_i)\phi(f_i)\right)
$$

여기서:

$$
n_i=\text{리뷰어 총 리뷰 수}
$$

$$
f_i=\text{리뷰어 팔로워 수}
$$

코드 변경 위치:

```python
def follower_weight(count: int) -> float:
    return reviewer_count_weight(count)
```

이 변경의 효과:

- 팔로워가 있는 리뷰어의 평가가 이전보다 더 크게 반영된다.
- 리뷰 수와 팔로워 수 모두 같은 누진 구조를 탄다.
- 그래도 최종 가중치 상한 $C=10$이 있어서, 팔로워가 아주 많은 리뷰어 한 명이 점수를 과도하게 지배하지는 못한다.
---

## 3.10 평균 별점 보너스 추가

모델을 확인하면서 다음 문제가 발견되었다.

> 리뷰어 평균 대비 점수와 조작 감점은 잘 반영되지만, 카카오에 표시되는 식당 평균 별점 자체가 순위에서 너무 약하게 반영된다.

기존 모델은 사실상 다음 구조였다.

$$
F_j=B_j-\lambda G_j
$$

여기서 $B_j$는 리뷰어 평균 대비 점수 기반이고, $G_j$는 조작 의심 점수다. 이 구조에서는 식당 평균 별점 $a_j$가 표에는 보이지만 최종 순위에는 거의 직접 반영되지 않았다.

그래서 평균 별점을 보너스 항으로 추가했다. 단, 평균 별점은 조작 가능성이 낮을수록 더 믿을 수 있으므로 $(1-G_j)$를 곱했다.

평균 별점 효과:

$$
E_j=\operatorname{clip}(a_j-4.0,-1,1)
$$

평균 별점 반영도:

$$
T_j=1-G_j
$$

평균 별점 보너스:

$$
R_j=\alpha T_jE_j
$$

현재 기본값:

$$
\alpha=0.35
$$

따라서 현재 최종 점수는 다음이다.

$$
F_j=B_j-0.5G_j+0.35(1-G_j)\operatorname{clip}(a_j-4.0,-1,1)
$$

이 설계의 의미:

- 평균 별점이 4.0보다 높으면 보너스를 받는다.
- 평균 별점이 4.0보다 낮으면 약한 패널티를 받는다.
- 조작의심점수 $G_j$가 높으면 평균 별점 보너스가 줄어든다.
- 조작의심점수 $G_j$가 낮으면 평균 별점이 더 강하게 반영된다.

코드 옵션:

```powershell
--rating-weight 0.35 --rating-baseline 4.0
```

코드상 기본값:

```python
parser.add_argument("--rating-weight", type=float, default=0.35, help="평균 별점 보너스 가중치")
parser.add_argument("--rating-baseline", type=float, default=4.0, help="평균 별점 보너스 기준점")
```

HTML/CSV에는 다음 컬럼을 추가했다.

| 컬럼 | 의미 |
|---|---|
| `평균별점효과` | $\operatorname{clip}(a_j-4.0,-1,1)$ |
| `평균별점반영도` | $1-G_j$ |
| `평균별점보너스` | $0.35(1-G_j)\operatorname{clip}(a_j-4.0,-1,1)$ |
| `조작후점수` | 평균 별점 보너스를 더하기 전 점수 |
| `최종보정점수` | 평균 별점 보너스까지 더한 최종 점수 |

---

---

## 3.11 총 리뷰수 기반 자동 스크롤 수집으로 변경

순위 모델을 계속 다듬는 과정에서, 분석 수식보다 먼저 데이터 수집 품질이 중요하다는 문제가 발견되었다.

기존 수집 방식은 모든 식당에 같은 스크롤 횟수를 적용했다.

예:

```powershell
--scrolls 10
```

이 방식의 문제:

- 리뷰가 20개인 식당은 충분히 수집될 수 있다.
- 리뷰가 300개인 식당은 10번 스크롤로 일부만 수집될 수 있다.
- 결국 `캡처된 리뷰수`가 `식당 총 리뷰수`보다 훨씬 적은 식당은 조작 의심 점수와 최종 점수가 불안정해진다.

그래서 `kakao_place_visible_text.mjs`에 총 리뷰수 기반 자동 스크롤을 추가했다. 이후 효율을 더 높이기 위해 총 리뷰수 40개 이하는 빠른 프로파일, 40개 초과는 느린 프로파일로 나누었다.

현재 기본 동작:

1. 페이지에 들어간 뒤 식당 총 리뷰수를 읽는다.
2. 현재 캡처된 리뷰 수를 파싱한다.
3. 총 리뷰수가 40개 이하이면 빠른 프로파일로 짧게 수집한다.
4. 총 리뷰수가 40개를 넘으면 느린 프로파일로 충분히 기다리며 수집한다.
5. 목표 수집률에 도달할 때까지 추가 스크롤한다.
6. 리뷰 수가 바로 늘지 않으면 추가 대기 후 다시 확인한다.
7. 추가 대기 후에도 리뷰 수가 반복해서 늘지 않으면 중단한다.
8. 최대 스크롤 횟수에 도달하면 중단한다.

추가된 옵션:

| 옵션 | 기본값 | 의미 |
|---|---:|---|
| `--no-auto-scroll` | 꺼짐 | 자동 스크롤을 끄고 기존처럼 `--scrolls`만 사용 |
| `--fast-threshold` | `40` | 총 리뷰수가 이 값 이하이면 빠른 프로파일 사용 |
| `--fast-max-scrolls` | `12` | 빠른 프로파일 최대 스크롤 횟수 |
| `--fast-stall-wait-ms` | `600` | 빠른 프로파일에서 리뷰 수가 안 늘 때 추가 대기 시간 |
| `--max-scrolls` | `200` | 느린 프로파일 자동 스크롤 최대 횟수 |
| `--target-coverage` | `0.95` | 총 리뷰수 대비 목표 수집률 |
| `--stall-limit` | `8` | 느린 프로파일에서 리뷰 수가 늘지 않는 스크롤이 몇 번 반복되면 멈출지 |
| `--stall-wait-ms` | `5000` | 느린 프로파일에서 리뷰 수가 안 늘 때 추가 대기 시간 |
| `--scroll-delay-ms` | `1200` | 느린 프로파일 스크롤 사이 대기 시간 |

수집 상태는 콘솔에 다음처럼 표시된다.

```text
리뷰 수집: 180/200 (90.0%), 스크롤 42회, 상태 max_scrolls_reached, 프로파일 slow
```

통합 XLSX의 `식당요약` 시트에도 다음 컬럼을 추가했다.

| 컬럼 | 의미 |
|---|---|
| `수집률` | `캡처된 리뷰수 / 식당 총 리뷰수` |
| `수집상태` | `target_reached`, `stalled`, `max_scrolls_reached`, `unknown_total`, `후기미제공` |
| `수집프로파일` | `fast`, `slow`, `fixed` 중 실제 사용한 수집 방식 |
| `스크롤횟수` | 실제 수행한 스크롤 횟수 |
| `목표리뷰수` | 목표 수집률 기준으로 필요한 리뷰 수 |

재수집 권장 명령:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --scrolls 3 --fast-threshold 40 --max-scrolls 200 --target-coverage 0.95 --stall-limit 8 --stall-wait-ms 5000 --wait-ms 2500 --scroll-delay-ms 1200 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx"
```

매우 꼼꼼히 모으고 싶을 때:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --scrolls 3 --fast-threshold 40 --max-scrolls 260 --target-coverage 1.0 --stall-limit 12 --stall-wait-ms 8000 --wait-ms 3000 --scroll-delay-ms 1500 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_전체시도.xlsx"
```

주의할 점:

- 카카오 페이지가 실제로 모든 리뷰를 무한 스크롤로 노출하지 않으면 `target_reached`에 도달하지 못할 수 있다.
- 매장주 요청으로 후기가 제공되지 않는 장소는 총 리뷰수와 수집상태를 `후기미제공`으로 표시한다.
- 이 경우 `max_scrolls_reached` 또는 `stalled`가 표시된다.
- 그래도 이전처럼 고정 10회 스크롤하는 방식보다 수집 품질을 훨씬 잘 확인할 수 있다.

# 4. 점수 계산 수식

## 4.1 리뷰별 기본 차이

리뷰어 $i$가 식당 $j$에 남긴 별점을 $r_{ij}$, 그 리뷰어의 평소 평균 별점을 $\bar r_i$라고 하자.

리뷰별 차이는 다음이다.

$$
d_{ij}=r_{ij}-\bar r_i
$$

의미:

- 어떤 리뷰어가 평소 평균 4.8점을 주는데 이 식당에 5점을 줬다면 차이는 $+0.2$다.
- 어떤 리뷰어가 평소 평균 3.2점을 주는데 이 식당에 5점을 줬다면 차이는 $+1.8$이다.

즉, 같은 5점이라도 까다로운 리뷰어가 준 5점을 더 가치 있게 본다.

---

## 4.2 리뷰어 신뢰 가중치

초기에는 로그 기반 가중치를 사용했다.

$$
w_i=
\min\left(
C,
(1+\ln(1+n_i))(1+\ln(1+f_i))
\right)
$$

여기서:

$$
n_i=\text{리뷰어의 총 리뷰 수}
$$

$$
f_i=\text{리뷰어의 팔로워 수}
$$

$$
C=\text{가중치 상한값}
$$

하지만 사용자가 지적한 문제는 이랬다.

> 리뷰 2개 쓴 사람과 10개 쓴 사람의 차이는 꽤 큰데, 로그를 바로 쓰면 초반 차이가 충분히 반영되지 않는다.

그래서 누진세처럼 구간별 가중치로 바꿨다.

---

## 4.3 현재 리뷰 수 가중치

리뷰어의 총 리뷰 수를 $n_i$라고 할 때:

$$
\phi(n_i)=
\begin{cases}
1+\dfrac{n_i}{3}, & 0\le n_i\le 10 \\
1+\dfrac{10}{3}+1.2\ln\left(\dfrac{n_i}{10}\right), & 10<n_i\le 100 \\
1+\dfrac{10}{3}+1.2\ln(10)+0.7\log_{10}\left(\dfrac{n_i}{100}\right), & n_i>100
\end{cases}
$$

의미:

- 리뷰 0개에서 10개까지는 거의 선형적으로 빠르게 신뢰도를 올린다.
- 10개 이후부터는 증가 속도를 줄인다.
- 100개 이후부터는 더 완만한 $\log_{10}$ 스케일을 쓴다.

직관적으로는:

| 리뷰 수 구간 | 해석 |
|---|---|
| 0~10개 | 리뷰어 경험 차이가 크게 반영됨 |
| 10~100개 | 경험이 많아질수록 가중치는 오르지만 완만해짐 |
| 100개 초과 | 매우 많은 리뷰 수의 영향력이 과도해지는 것을 막음 |

---

## 4.4 팔로워 가중치

팔로워 수를 $f_i$라고 할 때, 현재는 리뷰 수 가중치와 같은 누진 함수를 그대로 사용한다.

$$
\psi(f_i)=\phi(f_i)
$$

즉:

$$
\psi(f_i)=
\begin{cases}
1+\dfrac{f_i}{3}, & 0\le f_i\le 10 \\
1+\dfrac{10}{3}+1.2\ln\left(\dfrac{f_i}{10}\right), & 10<f_i\le 100 \\
1+\dfrac{10}{3}+1.2\ln(10)+0.7\log_{10}\left(\dfrac{f_i}{100}\right), & f_i>100
\end{cases}
$$

최종 리뷰어 가중치는:

$$
w_i=\min\left(C,\phi(n_i)\psi(f_i)\right)
$$

현재 기본 상한값은:

$$
C=10
$$

이 변경으로 팔로워가 있는 리뷰어의 영향력이 이전보다 더 강하게 반영된다. 다만 최종 가중치에는 상한 $C=10$이 있어서 과도한 영향력은 잘린다.

---

## 4.5 식당별 가중 평균 점수

식당 $j$의 유효 리뷰 집합을 $I_j$라고 하자.

$$
S_j=
\frac{
\sum_{i\in I_j} d_{ij}w_i
}{
\sum_{i\in I_j} w_i
}
$$

의미:

- 리뷰어 평균 대비 얼마나 높게 평가받았는지 계산한다.
- 리뷰 수와 팔로워를 고려한 가중 평균이다.

---

## 4.6 리뷰 수 신뢰도 보정

리뷰가 너무 적으면 한두 개 리뷰에 점수가 크게 흔들린다.

그래서 리뷰 수 $m_j$에 따라 신뢰도 계수 $c_j$를 곱한다.

$$
c_j=\frac{m_j}{m_j+k}
$$

현재 기본값:

$$
k=10
$$

조작 감점 전 점수는:

$$
B_j=S_jc_j
$$

예를 들어:

- 유효 리뷰가 1개면 $c_j=\frac{1}{11}$이라 점수가 크게 줄어든다.
- 유효 리뷰가 50개면 $c_j=\frac{50}{60}$이라 꽤 신뢰한다.

---

# 5. 조작 의심 점수 설계 변화

## 5.1 초기 아이디어

처음에는 단순히 특정 기간에 리뷰가 몰리는지를 보려고 했다.

7일 안에 가장 많이 몰린 리뷰 수:

$$
b_{7,j}=\max_{\text{7일 구간}}(\text{그 구간 안의 리뷰 수})
$$

전체 리뷰 대비 비율:

$$
q_{7,j}=\frac{b_{7,j}}{m_j}
$$

날짜 집중 위험:

$$
A_j=
\operatorname{clip}
\left(
\frac{q_{7,j}-0.18}{0.32},0,1
\right)
$$

여기서:

$$
\operatorname{clip}(x,0,1)=
\begin{cases}
0 & x<0 \\
x & 0\le x\le 1 \\
1 & x>1
\end{cases}
$$

해석:

- 7일 안에 전체 리뷰의 18% 이하만 몰리면 정상으로 본다.
- 50% 이상 몰리면 위험도가 1에 가까워진다.

---

## 5.2 문제점: 반복 조작을 못 잡음

사용자가 지적한 문제:

> 한 번에만 몰리는 게 아니라, 한 달 쉬고 7일 몰리고, 또 몇 달 뒤 7일 몰리는 식이면 단일 7일 최대값만으로는 부족하다.

그래서 다음 신호들을 추가했다.

- 버스트 묶음
- 버스트 밀도
- 버스트 주기성
- 리뷰 간격 균일성

---

## 5.3 버스트 묶음

리뷰 날짜를 정렬한 뒤, 서로 7일 이내로 이어진 리뷰들을 하나의 묶음으로 본다.

```text
2026-01-01, 2026-01-03, 2026-01-07 → 같은 묶음
2026-03-15 → 다음 묶음
```

그리고 묶음 안의 리뷰가 3개 이상이면 버스트 묶음으로 본다.

$$
\mathcal{C}_j=\{C: |C|\ge 3, \text{인접 리뷰 날짜 차이}\le 7\text{일}\}
$$

버스트 리뷰 수:

$$
b^{cluster}_j=\sum_{C\in\mathcal{C}_j}|C|
$$

버스트 비율:

$$
q^{cluster}_j=\frac{b^{cluster}_j}{m_j}
$$

버스트 밀도 위험:

$$
B^{density}_j=
\operatorname{clip}
\left(
\frac{q^{cluster}_j-0.25}{0.45},0,1
\right)
$$

---

## 5.4 버스트 주기성

버스트 묶음이 여러 개 있을 때, 각 묶음의 중심 날짜를 구한다.

묶음 $C$의 중심 날짜:

$$
\mu_C=\frac{1}{|C|}\sum_{t\in C}t
$$

인접한 버스트 중심 사이의 간격:

$$
g_l=\mu_{C_{l+1}}-\mu_{C_l}
$$

간격들의 변동계수:

$$
CV_g=\frac{\sigma(g)}{\mu(g)}
$$

주기성 위험:

$$
P_j=
\operatorname{clip}
\left(
\frac{0.35-CV_g}{0.35},0,1
\right)
$$

해석:

- 버스트 간격이 거의 일정하면 $CV_g$가 작다.
- $CV_g$가 작을수록 주기적으로 리뷰를 넣은 의심이 커진다.
- 버스트 묶음이 3개 이상일 때만 계산한다.

---

## 5.5 리뷰 간격 균일성

리뷰가 너무 기계적으로 일정한 간격으로 올라오는지도 본다.

고유 리뷰 날짜를 정렬한다.

$$
t_1<t_2<\cdots<t_k
$$

날짜 간격:

$$
h_l=t_{l+1}-t_l
$$

간격 변동계수:

$$
CV_h=\frac{\sigma(h)}{\mu(h)}
$$

균일 위험:

$$
R^{regular}_j=
\operatorname{clip}
\left(
\frac{0.45-CV_h}{0.45},0,1
\right)
$$

하지만 단순히 장사가 꾸준히 잘되는 식당도 일정하게 리뷰가 달릴 수 있다.

그래서 균일 위험은 버스트 밀도와 함께 있을 때만 강하게 반영하도록 했다.

$$
R^{regular*}_j=R^{regular}_jB^{density}_j
$$

---

## 5.6 저활동 고평점 위험

리뷰어가 거의 활동하지 않았고 팔로워도 없는데 높은 별점을 주는 경우를 본다.

$$
U_j=
\operatorname{clip}
\left(
\frac{u_j-0.25}{0.50},0,1
\right)
$$

여기서 $u_j$는 다음 조건을 만족하는 리뷰 비율이다.

$$
r_{ij}\ge 4.5,
\quad n_i\le 3,
\quad f_i=0
$$

다만 사용자가 지적한 것처럼:

> 활동이 적은 사람이 진짜 맛있어서 5점을 줬을 수도 있다.

그래서 이 항목의 가중치는 낮게 조정했다.

---

## 5.7 만점 성향 위험

평소에도 4.8점 이상을 주는 리뷰어가 이 식당에도 5점을 주는 경우는 정보량이 크지 않다.

$$
M_j=
\operatorname{clip}
\left(
\frac{p_j-0.40}{0.45},0,1
\right)
$$

여기서 $p_j$는 다음 조건을 만족하는 리뷰 비율이다.

$$
r_{ij}=5.0,
\quad \bar r_i\ge 4.8
$$

하지만 이것도 이미 $r_{ij}-\bar r_i$에서 어느 정도 반영된다.

그래서 최종 조작 의심 수식에서는 매우 낮은 가중치만 부여했다.

---

## 5.8 날짜 몰림을 별 5개 리뷰로 한정

마지막으로 중요한 수정이 있었다.

사용자 의견:

> 특정 시기에 몰렸는지 분류해야 하는 리뷰는 별 5개 리뷰로 한정하면 좋을 듯. 별 1개를 주게끔 주인이 조작할 이유는 없으니까.

그래서 날짜 관련 조작 신호는 전부 별 5개 리뷰만 대상으로 바꿨다.

코드 핵심:

```python
five_star_valid = [item for item in valid if item["given"] >= 5.0 and item.get("date")]
date_n = len(five_star_valid)
dates = [item["date"] for item in five_star_valid]
```

이 변경의 의미:

- 품질 점수 $S_j$는 전체 유효 리뷰를 사용한다.
- 조작 의심의 날짜/버스트/주기성 신호는 별 5개 리뷰만 사용한다.
- 별 1~3개 리뷰가 특정 시기에 몰린 것을 홍보 조작으로 오해하지 않는다.

HTML 표에는 `별5유효` 컬럼을 추가했다.

---

# 6. 현재 최종 조작 의심 점수

현재 조작 의심 점수는 다음이다.

$$
G_j=
\operatorname{clip}
\left(
0.25A_j
+0.15D_j
+0.25B_j
+0.15P_j
+0.10R^{regular*}_j
+0.06U_j
+0.04M_j,
0,1
\right)
$$

각 항목:

| 기호 | 이름 | 의미 | 현재 계산 대상 |
|---|---|---|---|
| $A_j$ | 날짜집중위험 | 별 5개 리뷰가 가장 빽빽한 7일 구간에 얼마나 몰렸는가 | 별 5개 리뷰 |
| $D_j$ | 동일일자위험 | 별 5개 리뷰가 같은 날짜에 얼마나 몰렸는가 | 별 5개 리뷰 |
| $B_j$ | 버스트밀도위험 | 별 5개 리뷰가 3개 이상씩 붙은 묶음에 얼마나 많이 들어가는가 | 별 5개 리뷰 |
| $P_j$ | 버스트주기성위험 | 버스트 묶음들이 규칙적 간격으로 반복되는가 | 별 5개 리뷰 |
| $R^{regular*}_j$ | 보정리뷰간격균일위험 | 리뷰 간격이 균일하면서 버스트 밀도도 높은가 | 별 5개 리뷰 |
| $U_j$ | 저활동고평점위험 | 활동 적은 리뷰어의 고평점 비율 | 전체 유효 리뷰 |
| $M_j$ | 만점성향위험 | 원래 후한 리뷰어의 5점 비율 | 전체 유효 리뷰 |

날짜 계열 항목을 별 5개 리뷰로 제한한 이유:

- 주인이 조작한다면 보통 고평점 리뷰를 넣을 가능성이 높다.
- 낮은 별점 리뷰가 몰리는 것은 광고 조작보다는 사건, 서비스 문제, 외부 이슈일 가능성이 있다.
- 따라서 별 1~3개 리뷰까지 날짜 몰림에 넣으면 조작 의심 점수가 왜곡될 수 있다.

---

## 7. 조작 감점과 최종 점수

조작 감점:

$$
H_j=\lambda G_j
$$

현재 기본값:

$$
\lambda=0.5
$$

최종 보정 점수:

$$
F_j=B_j-H_j
$$

즉:

$$
F_j=S_jc_j-\lambda G_j
$$

해석:

- $S_j$: 리뷰어 평균 대비 얼마나 좋은 평가를 받았는가
- $c_j$: 리뷰 수가 충분한가
- $G_j$: 조작 의심 신호가 있는가
- $\lambda$: 조작 의심 점수를 실제 점수에서 얼마나 강하게 깎을 것인가
- $a_j$: 카카오에 표시되는 식당 평균 별점
- 평균 별점 보너스는 조작 의심이 낮을수록 강하게 반영된다

---

## 8. HTML 순위표

최종 HTML 생성 파일 예:

```text
outputs/kakao_place_text/송도과학로_순위표_별5집중보정.html
```

표에 포함되는 주요 컬럼:

| 컬럼 | 의미 |
|---|---|
| 순위 | 최종보정점수 기준 순위 |
| 식당 | 식당명과 카테고리 |
| 최종 | 최종보정점수 $F_j$ |
| 조작전 | 조작 감점 전 점수 $B_j$ |
| 의심 | 조작의심점수 $G_j$ |
| 감점 | 조작감점 $H_j$ |
| 원점수 | 리뷰어 평균 대비 점수 $S_j$ |
| 신뢰도 | 리뷰 수 신뢰도 $c_j$ |
| 식당평균 | 카카오 표시 평균 별점 |
| 총리뷰 | 카카오 표시 총 리뷰수 |
| 캡처 | 실제 캡처된 리뷰수 |
| 유효 | 점수 계산에 사용된 리뷰수 |
| 최대7일 | 별 5개 리뷰 중 7일 안에 가장 많이 몰린 수 |
| 동일일 | 별 5개 리뷰 중 같은 날짜 최대 수 |
| 별5유효 | 날짜 몰림 계산에 사용한 별 5개 리뷰 수 |
| 버스트묶음 | 별 5개 리뷰 버스트 묶음 수 |
| 버스트비율 | 별 5개 리뷰 중 버스트에 들어간 비율 |
| 버스트밀도 | 버스트 밀도 위험 |
| 주기성 | 버스트 주기성 위험 |
| 균일보정 | 간격 균일성 × 버스트 밀도 |
| 저활동고평점 | 저활동 리뷰어 고평점 위험 |

---

## 9. 생성된 주요 결과 파일

### 음식점 목록

```text
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx
```

### 통합 리뷰

```text
outputs/kakao_place_text/송도과학로_통합리뷰.xlsx
```

### 순위표 변천

| 파일 | 의미 |
|---|---|
| `송도과학로_순위표.html` | 초기 순위표 |
| `송도과학로_순위표_조작보정.html` | 조작 감점 추가 버전 |
| `송도과학로_순위표_누진가중치.html` | 리뷰어 가중치를 누진식으로 변경 |
| `송도과학로_순위표_누진가중치_날짜중심.html` | 날짜 집중 위험 가중치를 중심으로 조정 |
| `송도과학로_순위표_주기보정.html` | 반복 버스트/주기성 보정 추가 |
| `송도과학로_순위표_별5집중보정.html` | 날짜/버스트 위험을 별 5개 리뷰로 한정한 현재 버전 |

---

## 10. 현재 재실행 명령 모음

## 10.1 음식점 목록 수집

```powershell
$env:KAKAO_REST_API_KEY="REST_API_KEY_값"
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도과학로"
```

## 10.2 리뷰 통합 XLSX 생성

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --scrolls 3 --fast-threshold 40 --max-scrolls 200 --target-coverage 0.95 --stall-limit 8 --stall-wait-ms 5000 --wait-ms 2500 --scroll-delay-ms 1200 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx"
```

테스트용 10개:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --limit 10 --scrolls 10 --wait-ms 1500 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_10개.xlsx"
```

## 10.3 순위표 생성

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_별5집중보정.html"
```

옵션 조절 예:

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰.xlsx" --k 10 --weight-cap 10 --penalty-scale 0.5 --rating-weight 0.35 --rating-baseline 4.0 --out "outputs\kakao_place_text\송도과학로_순위표_별5집중보정.html"
```

---

## 11. 코드별 역할 요약

## 11.1 `kakao_restaurant_search.py`

역할:

- Kakao Local API 호출
- 지역 기반 음식점 목록 수집
- 리뷰 URL 생성
- JSON/CSV/XLSX 저장

핵심 설계:

- `keyword` 검색과 `category` 검색을 모두 지원한다.
- `category` 검색은 음식점 코드 `FD6`을 사용한다.
- 중복 장소는 카카오 장소 ID 기준으로 제거한다.
- `openpyxl`이 없으면 XLSX 저장만 건너뛰고 CSV/JSON은 저장한다.

---

## 11.2 `kakao_place_visible_text.mjs`

역할:

- 카카오 플레이스 리뷰 페이지 접속
- 스크롤로 리뷰 추가 로드
- visible text 저장
- TXT/JSON/XLSX 생성
- 여러 URL을 통합 XLSX로 저장

핵심 설계:

- `#review` URL로 정규화한다.
- XLSX의 `review_url` 열을 읽는다.
- 리뷰어 닉네임은 저장하지 않고 `익명1`, `익명2` 형태로 저장한다.
- `--combined-xlsx`가 있으면 식당별 XLSX 대신 하나의 통합 XLSX를 만든다.

---

## 11.3 `kakao_review_ranker.py`

역할:

- 통합 리뷰 XLSX 읽기
- 식당별 리뷰 그룹화
- 리뷰어 평균 대비 점수 계산
- 리뷰어 신뢰 가중치 계산
- 리뷰 수 신뢰도 보정
- 조작 의심 점수 계산
- HTML/CSV 순위표 생성

핵심 설계:

- 외부 패키지 없이 XLSX ZIP/XML을 직접 읽는다.
- 품질 점수는 전체 유효 리뷰를 사용한다.
- 날짜/버스트/주기성 조작 위험은 별 5개 리뷰만 사용한다.
- HTML에는 수식 설명과 주요 위험 지표를 같이 보여준다.

---

## 12. 한계와 주의점

이 분석은 “확률적 의심 점수”이지, 조작 여부를 확정하는 도구는 아니다.

주의할 점:

- 카카오 페이지 구조가 바뀌면 visible text 파싱이 깨질 수 있다.
- 스크롤 횟수가 부족하면 일부 리뷰만 수집된다.
- 카카오가 비로그인/자동화 접근에 제한을 걸면 수집량이 줄 수 있다.
- 별 5개가 몰렸다고 무조건 조작은 아니다. 이벤트, 방송 노출, 신규 오픈, 메뉴 출시 등도 원인이 될 수 있다.
- 반대로 조작이 아주 천천히 섞이면 현재 지표로도 완전히 잡기 어렵다.

그래서 최종 순위는 다음처럼 읽는 것이 좋다.

- `최종보정점수`가 높고 `조작의심점수`가 낮으면 비교적 좋은 후보
- `조작전점수`는 높은데 `조작의심점수`도 높으면 직접 리뷰 내용을 확인할 필요 있음
- `별5유효`, `버스트비율`, `최대7일`, `동일일`을 함께 보면 왜 감점됐는지 이해하기 쉬움

---

## 13. 다음에 추가하면 좋은 기능

후속 개선 후보:

1. 리뷰 텍스트 감성 분석 추가
2. 같은 문구가 반복되는 리뷰 탐지
3. 리뷰어들이 여러 식당에 같은 날짜에 같은 패턴으로 리뷰했는지 탐지
4. 신규 오픈일이나 이벤트 기간을 반영한 예외 처리
5. 식당 카테고리별 평균 점수 정규화
6. 리뷰 수가 너무 적은 식당은 별도 “판단 보류” 그룹으로 분리
7. HTML에서 조작 의심 높은 항목만 필터링하는 기능 추가
8. 날짜 분포 히스토그램을 HTML에 시각화

---

## 14. 현재 최종 기준 요약

현재 가장 최신 기준은 다음이다.

$$
F_j=S_jc_j-\lambda G_j
$$

$$
S_j=
\frac{
\sum_{i\in I_j}(r_{ij}-\bar r_i)w_i
}{
\sum_{i\in I_j}w_i
}
$$

$$
c_j=\frac{m_j}{m_j+10}
$$

$$
\psi(f_i)=\phi(f_i)
$$

$$
w_i=\min\left(10,\phi(n_i)\phi(f_i)\right)
$$

$$
G_j=
\operatorname{clip}
\left(
0.25A_j
+0.15D_j
+0.25B_j
+0.15P_j
+0.10R^{regular*}_j
+0.06U_j
+0.04M_j,
0,1
\right)
$$

$$
\lambda=0.5
$$

최신 해석:

> 좋은 식당은 평소보다 후하게 평가받고 평균 별점도 높은 식당이며, 별 5개 리뷰가 특정 시기에 비정상적으로 몰린 흔적이 있으면 평균 별점의 영향도 줄이고 점수도 깎는다.












