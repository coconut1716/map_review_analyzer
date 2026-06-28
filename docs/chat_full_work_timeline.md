# 전체 작업 기록: 시간표, 연세 공간대관, 카카오맵 식당 리뷰 분석

작성일: 2026-06-26  
프로젝트 위치: `C:\Users\kimwo\project_review`  
이전 작업 위치: `C:\Users\kimwo\OneDrive\Desktop\대학 교재\반품향`

이 문서는 이 채팅에서 처음부터 끝까지 진행한 작업을 한 번에 다시 이해할 수 있도록 정리한 기록이다. 핵심은 세 갈래였다.

1. 연세대 강의/공간대관/시간표 관련 HTML과 XLSX 처리
2. 카카오맵 식당 목록 및 리뷰 수집 자동화
3. 리뷰어 신뢰도, 리뷰 조작 의심, 실제 리뷰 텍스트까지 반영한 식당 순위표 생성

---

## 1. 전체 흐름 요약

처음에는 HTML 시간표와 강의 정보 표시를 다루다가, 연세대 공간대관 시스템에서 대관/강의 정보를 어떻게 뽑을지로 확장했다. 이후 카카오맵 음식점 데이터를 API와 Playwright로 수집하고, 단순 별점이 아니라 리뷰어 성향, 리뷰 날짜 집중, 조작 의심, 실제 리뷰 문구까지 반영하는 식당 순위 모델을 만들었다.

최종적으로 프로젝트는 OneDrive 동기화 문제를 피하기 위해 다음 위치로 옮겼다.

```powershell
C:\Users\kimwo\project_review
```

이 위치에서는 OneDrive 동기화 때문에 생기는 느림, 파일 잠금, XLSX 저장 지연 문제가 줄어든다.

---

## 2. 주요 폴더 구조

```text
project_review/
  docs/
    kakao_restaurant_analysis_timeline.md
    chat_full_work_timeline.md
  xlsx_build/
    kakao_restaurant_search.py
    kakao_place_visible_text.mjs
    kakao_review_ranker.py
    kakao_review_text_ranker.py
    merge_retry_reviews.mjs
    yonsei_space_export_from_curl.mjs
    build_timetable_html.mjs
    timetable_client.js
    timetable_style.css
    grade_distribution_to_html.py
  outputs/
    kakao_restaurants/
    kakao_place_text/
  node_modules/
  package.json
```

중요한 점:

- `node_modules`는 같이 복사해 두는 것이 좋다.
- `package.json`에는 `playwright-core`만 기록되어 있지만, XLSX 생성 쪽은 `@oai/artifact-tool`도 사용한다.
- 그래서 새 폴더에서 단순히 `npm install`만 하면 일부 XLSX 생성 스크립트가 깨질 수 있다.

---

# Part A. 연세대 강의/시간표/공간대관 작업

## 3. HTML 시간표 수정 요청

초기 요청은 기존 HTML 파일을 바로 수정하지 않고, 복사본을 만들어 뒤에 `modify`를 붙인 파일을 만든 뒤, 시간표 카드에 기존 표시 항목인 과목명, 시간, 강의실 아래에 교수명까지 넣는 것이었다.

핵심 요구:

- 원본 HTML은 건드리지 않는다.
- 복사본을 만든다.
- 시간표 블록에는 다음 정보가 보여야 한다.

```text
과목명
시간
강의실
교수명
```

이 흐름은 이후 Excel 기반 시간표 생성기 작업으로 이어졌다.

---

## 4. 연세대학교 공간대관 시스템 조사

대상 사이트:

```text
https://space.yonsei.ac.kr/index.php?mid=K06
```

처음에는 사용자가 직접 로그인한 상태에서 Chrome DevTools Network 탭을 열고, 대관현황 조회 요청을 cURL로 복사하는 방법을 찾았다.

우리가 확인한 점:

- 단순히 페이지 URL만으로는 데이터가 바로 드러나지 않는다.
- 로그인 세션과 요청 파라미터가 필요하다.
- DevTools의 Network 탭에서 Fetch/XHR 요청을 잡고 cURL로 복사하면, 그 요청을 재현할 수 있다.
- 공간대관 상세 팝업에는 `주관단체명`에 학과나 교수명 성격의 정보가 들어갈 수 있다.

---

## 5. 연세 공간대관 cURL 기반 추출기

파일:

```text
xlsx_build/yonsei_space_export_from_curl.mjs
```

역할:

- DevTools에서 복사한 cURL 요청을 읽는다.
- 쿠키와 헤더를 재사용한다.
- 대관현황 페이지의 AJAX 요청을 재현한다.
- 국제캠퍼스, 건물, 강의실, 날짜 범위를 돌면서 수업/대관 정보를 추출한다.
- 결과를 JSON/XLSX 형태로 저장한다.

주요 대상 기간:

```text
2026-09-07 ~ 2026-09-11
```

사용 목적:

- 국제캠퍼스 전체 건물/강의실의 강의 정보를 정리한다.
- 학정번호, 과목명, 교수명, 시간, 강의실 정보를 시간표 생성에 활용한다.

---

## 6. 시간표 HTML 생성기

파일:

```text
xlsx_build/build_timetable_html.mjs
xlsx_build/timetable_client.js
xlsx_build/timetable_style.css
```

목표:

- 수집한 강의/대관 정보를 바탕으로 모바일 친화적인 시간표 HTML을 만든다.
- 사용자가 보낸 UI 이미지처럼 강의 리스트와 주간 시간표를 구성한다.

UI 요구사항:

- 캠퍼스는 국제캠퍼스만 사용한다.
- 전공/영역은 학정번호 앞 영어 prefix 기준으로 분류한다.
- 검색창은 2개를 둔다.

```text
학정번호 검색
과목명 검색
```

시간표 카드에는 다음 정보가 들어간다.

```text
과목명
교수명
강의실
```

---

## 7. 전공과목 성적 분포 XLSX HTML화

파일:

```text
xlsx_build/grade_distribution_to_html.py
```

입력 예:

```text
전공과목 성적 분포 (대학)_2026-05-0892138819.xlsx
```

요구사항:

- 1학기와 2학기를 나눈다.
- 사용자가 보낸 이미지처럼 표 형태로 정렬한다.
- 2025 데이터는 제외한다.
- A비율 기준 정렬 표를 HTML로 만든다.

결과적으로 성적 분포 XLSX를 보기 쉬운 HTML 표로 변환하는 흐름을 만들었다.

---

# Part B. 카카오맵 음식점 수집

## 8. 카카오맵 직접 스크래핑 가능성 검토

처음에는 사용자가 카카오맵에서 특정 지역 음식점을 검색하고, 검색 결과를 스크롤하며 음식점 이름과 리뷰 링크를 모으는 방식이 가능한지 물었다.

확인한 점:

- 카카오맵은 지도 내부에서 검색해도 주소창 URL이 잘 바뀌지 않는다.
- 검색 결과는 프론트엔드 내부 상태와 동적 DOM으로 관리된다.
- UI 직접 스크래핑은 가능하지만 느리고 깨지기 쉽다.
- 음식점 목록 수집은 Kakao Local REST API를 쓰는 것이 훨씬 안정적이다.

그래서 목록 수집은 API, 상세 리뷰 수집은 Playwright 기반 페이지 접근으로 나누었다.

---

## 9. Kakao Developers REST API 키 발급

사용자가 Kakao Developers에서 앱을 만들고 REST API 키를 발급받았다.

진행 중 확인한 점:

- 앱 생성 시 도메인은 웹 서비스용 설정에서 필요할 수 있지만, Local REST API 호출 자체에는 REST API 키가 핵심이다.
- 오류 `disabled OPEN_MAP_AND_LOCAL service`가 발생했다.
- 해결은 Kakao Developers 앱 설정에서 `지도/로컬` 또는 `OPEN_MAP_AND_LOCAL` 서비스를 활성화하는 것이었다.

PowerShell에서 키 설정:

```powershell
$env:KAKAO_REST_API_KEY="REST_API_KEY_값"
```

주의:

- PowerShell에서는 성공해도 별도 성공 메시지가 뜨지 않는다.
- 에러가 없으면 설정된 것이다.
- 경로에 공백이 있으면 반드시 따옴표를 써야 한다.

```powershell
cd "C:\Users\kimwo\project_review"
```

---

## 10. 음식점 목록 수집기

파일:

```text
xlsx_build/kakao_restaurant_search.py
```

역할:

- Kakao Local REST API를 호출한다.
- 지역명 + 음식점 키워드로 검색한다.
- 음식점 category 검색도 함께 사용한다.
- 중복 장소는 카카오 place id 기준으로 제거한다.
- `place_url`과 `review_url`을 생성한다.
- JSON, CSV, XLSX를 저장한다.

대표 실행:

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도동"
```

또는:

```powershell
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도과학로"
```

주요 결과 파일:

```text
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도동_20260623_133656.xlsx
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx
```

---

## 11. 송도과학로 검색의 한계

우리가 `송도과학로`로 검색해 음식점 URL을 뽑았지만, 모든 주변 식당이 잡히는 것은 아니었다.

예시:

```text
https://place.map.kakao.com/1922201261#review
```

이 URL은 확인 결과 다음 식당이었다.

```text
장소명: 덕이네
카테고리: 한식
카카오 표시 문구: 연수구 칼국수 인기 맛집
```

이 식당은 `송도과학로`보다 다음 검색어로 찾는 것이 더 적절하다.

```text
송도 덕이네
인천 연수구 덕이네
```

즉, 도로명 검색은 빠질 수 있으므로, 실제 완전성을 높이려면 다음 검색을 병행하는 편이 좋다.

```text
송도과학로 음식점
송도동 음식점
인천 연수구 송도동 한식
인천 연수구 송도동 카페
송도 트리플스트리트 음식점
식당명 직접 검색
```

---

# Part C. 카카오 플레이스 리뷰 수집

## 12. 카카오 플레이스 visible text 수집기

파일:

```text
xlsx_build/kakao_place_visible_text.mjs
```

역할:

- `https://place.map.kakao.com/{placeId}#review` URL에 접속한다.
- Playwright로 페이지를 연다.
- 브라우저에 보이는 텍스트를 읽는다.
- 리뷰어 이름은 저장하지 않고 `익명1`, `익명2` 식으로 익명화한다.
- 리뷰 날짜, 리뷰어 평균 별점, 이 식당에 준 별점, 리뷰어 총 리뷰수, 팔로워 수, 리뷰 내용을 파싱한다.
- TXT, JSON, XLSX를 저장한다.
- 여러 식당을 하나의 통합 XLSX로 저장할 수 있다.

단일 URL 실행 예:

```powershell
node xlsx_build\kakao_place_visible_text.mjs "https://place.map.kakao.com/450422239#review" --scrolls 10
```

XLSX의 모든 `review_url` 실행 예:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰.xlsx"
```

---

## 13. 통합 리뷰 XLSX 구조

대표 파일:

```text
outputs/kakao_place_text/송도과학로_통합리뷰_자동스크롤.xlsx
```

시트:

```text
식당요약
리뷰전체
```

`식당요약` 주요 컬럼:

```text
식당번호
장소명
카테고리
식당 평균 별점
식당 총 리뷰수
캡처된 리뷰수
수집률
수집상태
수집프로파일
스크롤횟수
목표리뷰수
블로그 수
원본 URL
원본 TXT
```

`리뷰전체` 주요 컬럼:

```text
식당번호
장소명
식당 평균 별점
식당 총 리뷰수
익명 리뷰어
리뷰 날짜
리뷰어 평균 별점
이 식당에 준 별점
리뷰어 총 리뷰수
팔로워 수
리뷰 내용
원본 URL
```

---

## 14. 자동 스크롤 개선

처음에는 모든 식당에 같은 스크롤 횟수를 적용했다. 하지만 리뷰가 적은 식당과 많은 식당을 같은 방식으로 처리하면 비효율적이었다.

그래서 자동 스크롤을 개선했다.

현재 방식:

- 총 리뷰수 40개 이하: 빠른 프로파일 `fast`
- 총 리뷰수 40개 초과: 느린 프로파일 `slow`
- 자동 스크롤 끔: 고정 프로파일 `fixed`

주요 옵션:

```text
--fast-threshold 40
--fast-max-scrolls 12
--fast-stall-wait-ms 600
--max-scrolls 200
--target-coverage 0.95
--stall-limit 8
--stall-wait-ms 5000
--scroll-delay-ms 1200
```

현재 권장 실행:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --scrolls 3 --fast-threshold 40 --max-scrolls 200 --target-coverage 0.95 --stall-limit 8 --stall-wait-ms 5000 --wait-ms 2500 --scroll-delay-ms 1200 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx"
```

---

## 15. 후기미제공 처리

일부 식당은 실제로 카카오에서 후기를 제공하지 않는다.

처리 방식:

```text
식당 총 리뷰수: 후기미제공
수집상태: 후기미제공
캡처된 리뷰수: 0
```

이 케이스는 실패가 아니라 정상적인 수집 불가 상태로 보았다.

---

## 16. 3개 식당 재수집

자동 수집 후 실제 리뷰가 있는데 수집이 덜 된 식당 3개를 확인했다.

```text
65/89 https://place.map.kakao.com/1684444798#review  스시사쿠 송도점
73/89 https://place.map.kakao.com/300216955#review   정씨함박 송도점
87/89 https://place.map.kakao.com/1733825326#review  한끼갈비탕 송도 트리플스트리트 본점
```

재수집 URL 파일:

```text
outputs/kakao_place_text/retry_failed_3_urls.txt
```

재수집 명령:

```powershell
node xlsx_build\kakao_place_visible_text.mjs --url-file "outputs\kakao_place_text\retry_failed_3_urls.txt" --scrolls 3 --fast-threshold 0 --max-scrolls 300 --target-coverage 1.0 --stall-limit 15 --stall-wait-ms 10000 --wait-ms 3500 --scroll-delay-ms 1800 --combined-xlsx "outputs\kakao_place_text\재수집_실패3개.xlsx"
```

재수집 결과:

```text
스시사쿠 송도점: 156개
정씨함박 송도점: 213개
한끼갈비탕 송도 트리플스트리트 본점: 92개
```

---

## 17. 재수집 결과 병합

파일:

```text
xlsx_build/merge_retry_reviews.mjs
```

역할:

- 기존 통합 XLSX에서 위 3개 URL의 요약 행을 재수집본으로 교체한다.
- 기존 통합 XLSX에서 위 3개 URL의 리뷰 행을 삭제한다.
- 재수집 XLSX의 리뷰 행을 다시 삽입한다.
- 식당번호는 기존 번호를 유지한다.

병합 후 결과:

```text
전체 리뷰 행: 1539개 -> 1880개
식당 수: 89개 유지
```

최종 통합 파일은 사용자가 이름을 다시 맞춰 두었다.

```text
outputs/kakao_place_text/송도과학로_통합리뷰_자동스크롤.xlsx
```

---

# Part D. 숫자 기반 식당 순위 모델

## 18. 기본 순위표 생성기

파일:

```text
xlsx_build/kakao_review_ranker.py
```

역할:

- 통합 리뷰 XLSX를 읽는다.
- 식당별 리뷰를 묶는다.
- 리뷰어 평균 별점과 해당 식당 별점의 차이를 계산한다.
- 리뷰어 신뢰 가중치를 계산한다.
- 리뷰 수 신뢰도 보정을 적용한다.
- 리뷰 날짜 집중과 조작 의심 점수를 계산한다.
- 평균 별점 보너스를 넣는다.
- 최종 HTML/CSV 순위표를 만든다.

실행:

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_자동스크롤_현재파일.html"
```

---

## 19. 리뷰별 기본 차이

리뷰어 $i$가 식당 $j$에 남긴 별점을 $r_{ij}$, 그 리뷰어의 평소 평균 별점을 $\bar r_i$라고 하면:

$$
d_{ij}=r_{ij}-\bar r_i
$$

해석:

- 평소 4.8점을 주는 리뷰어가 5점을 주면 차이는 작다.
- 평소 3.2점을 주는 리뷰어가 5점을 주면 차이는 크다.
- 같은 5점이라도 까다로운 사람이 준 5점을 더 높게 본다.

---

## 20. 리뷰어 신뢰 가중치

처음에는 로그 가중치만 쓰려 했지만, 리뷰 2개와 10개의 차이를 더 크게 봐야 한다는 판단으로 누진식으로 바꿨다.

리뷰 수 가중치:

$$
\phi(n)=
\begin{cases}
1+\dfrac{n}{3}, & 0\le n\le 10 \\
1+\dfrac{10}{3}+1.2\ln\left(\dfrac{n}{10}\right), & 10<n\le 100 \\
1+\dfrac{10}{3}+1.2\ln(10)+0.7\log_{10}\left(\dfrac{n}{100}\right), & n>100
\end{cases}
$$

팔로워 가중치는 리뷰 수 가중치와 동일하게 바꿨다.

$$
\psi(f)=\phi(f)
$$

최종 리뷰어 가중치:

$$
w_i=\min\left(10,\phi(n_i)\psi(f_i)\right)
$$

---

## 21. 식당별 가중 평균 점수

$$
S_j=
\frac{
\sum_{i\in I_j}d_{ij}w_i
}{
\sum_{i\in I_j}w_i
}
$$

리뷰 수 신뢰도:

$$
c_j=\frac{m_j}{m_j+10}
$$

조작 감점 전 점수:

$$
B_j=S_jc_j
$$

---

## 22. 조작 의심 점수

초기에는 7일 날짜 집중만 보려 했지만, 반복적인 조작을 잡기 위해 여러 신호를 추가했다.

현재 조작 의심 점수:

$$
G_j=
\operatorname{clip}
\left(
0.25A_j
+0.15D_j
+0.25B^{density}_j
+0.15P_j
+0.10R^{regular*}_j
+0.06U_j
+0.04M_j,
0,1
\right)
$$

각 항목:

| 기호 | 의미 |
|---|---|
| $A_j$ | 별 5개 리뷰의 7일 날짜 집중 위험 |
| $D_j$ | 별 5개 리뷰의 동일 날짜 집중 위험 |
| $B^{density}_j$ | 별 5개 리뷰 버스트 묶음 밀도 |
| $P_j$ | 버스트 묶음의 주기성 위험 |
| $R^{regular*}_j$ | 리뷰 간격 균일성 × 버스트 밀도 |
| $U_j$ | 저활동 리뷰어 고평점 위험 |
| $M_j$ | 원래 후한 리뷰어의 만점 성향 위험 |

중요 수정:

- 날짜 집중, 버스트, 주기성 위험은 별 5개 리뷰만 대상으로 계산한다.
- 별 1개 리뷰가 몰린 것을 주인이 조작했다고 보기는 어렵기 때문이다.

---

## 23. 조작 감점과 평균 별점 반영

조작 감점:

$$
H_j=0.5G_j
$$

기존에는 최대 감점 계수가 0.35였으나, 너무 후하다고 판단해 0.5로 높였다.

조작 후 점수:

$$
Q_j=B_j-H_j
$$

평균 별점 효과:

$$
E_j=\operatorname{clip}(a_j-4.0,-1,1)
$$

평균 별점 보너스:

$$
R_j=0.35(1-G_j)E_j
$$

최종 보정 점수:

$$
F_j=Q_j+R_j
$$

해석:

- 조작 의심이 낮을수록 평균 별점의 영향력을 더 믿는다.
- 조작 의심이 높으면 평균 별점 보너스가 줄어든다.

---

## 24. 숫자 기반 순위 결과

최신 통합 파일 기준으로 숫자 기반 순위표를 만들었다.

출력:

```text
outputs/kakao_place_text/송도과학로_순위표_자동스크롤_현재파일.html
outputs/kakao_place_text/송도과학로_순위표_자동스크롤_현재파일.csv
```

상위권 예:

| 순위 | 식당 | 특징 |
|---:|---|---|
| 1 | 선율페이스트리 | 높은 평균별점, 조작 의심 낮음 |
| 2 | 파티세리키세츠 | 평균별점 5.0, 리뷰 수는 적지만 긍정적 |
| 3 | 후다닥냉면 | 리뷰 수와 평균별점 균형 좋음 |
| 4 | 프레피커피랩 | 조작 의심 낮고 평균별점 높음 |
| 5 | 후니네김밥 | 리뷰 수는 적지만 평가 좋음 |
| 7 | 스시사쿠 송도점 | 재수집 후 156개 리뷰 반영 |

조작 의심이 높게 나온 식당 예:

```text
캡틴루이
뽁식당 송도점
정씨함박 송도점
샤브혜
한끼갈비탕 송도 트리플스트리트 본점
```

이것은 맛없다는 뜻이 아니라, 별 5개 리뷰가 날짜/버스트 패턴상 의심스럽다는 뜻이다.

---

# Part E. 실제 리뷰 내용 반영 순위표

## 25. 리뷰 텍스트 반영 순위 모델

파일:

```text
xlsx_build/kakao_review_text_ranker.py
```

요구:

> 실제 리뷰 내용까지 모두 읽고 순위표를 처음부터 끝까지 다시 매겨줘. 리뷰 알바, 조작 이런 말이 있으면 순위를 내려줘.

그래서 기존 숫자 기반 점수에 리뷰 텍스트 신호를 추가했다.

읽은 리뷰 텍스트 수:

```text
1672개
```

---

## 26. 텍스트 감점 신호

직접 조작 언급:

```text
리뷰 알바
별점 알바
평점 알바
가짜 리뷰
리뷰 조작
별점 조작
알바 의심
```

광고/협찬/이벤트 언급:

```text
체험단
협찬
제공받음
원고료
리뷰 이벤트로 점수 뻥튀기
광고 ㄹㅈㄷ
```

주의:

초기에는 `광고`나 `리뷰 이벤트`를 너무 넓게 잡아서 거짓 양성이 있었다.

예:

```text
냉면 광고 입간판
리뷰 이벤트 말고 제 의지로 씀
```

이런 문장은 조작 신호가 아니므로 패턴을 더 엄격하게 고쳤다.

강한 불만:

```text
최악
다신 안 감
재방문 안 함
비추
맛없음
불친절
위생
머리카락
벌레
식중독
환불
```

일반 불만:

```text
별로
실망
아쉽
비싸
늦음
오래 걸림
딱딱함
질김
냄새
불편
```

긍정 신호:

```text
맛있다
맛집
친절
추천
재방문
또 갈
깔끔
신선
만족
가성비
푸짐
```

---

## 27. 내용 반영 점수

기존 최종 보정 점수를 $F_j$라고 하면:

$$
T_j=F_j-\text{텍스트감점}+\text{텍스트보너스}
$$

텍스트 감점은 다음을 반영한다.

- 조작/알바 직접 언급은 크게 감점
- 리뷰 이벤트/광고 의심은 중간 감점
- 강한 불만은 감점
- 일반 불만은 약한 감점

텍스트 보너스는 긍정 신호가 불만 신호보다 많을 때만 작게 반영한다. 긍정 문구 남발로 점수가 과하게 오르지 않도록 상한을 낮게 두었다.

---

## 28. 내용 반영 순위표 생성

명령:

```powershell
python xlsx_build\kakao_review_text_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_리뷰내용반영.html"
```

출력:

```text
outputs/kakao_place_text/송도과학로_순위표_리뷰내용반영.html
outputs/kakao_place_text/송도과학로_순위표_리뷰내용반영.csv
```

최종 상위권 예:

| 순위 | 식당 | 내용반영점수 | 특징 |
|---:|---|---:|---|
| 1 | 선율페이스트리 | 0.9191 | 긍정 신호 많고 감점 없음 |
| 2 | 후다닥냉면 | 0.8545 | 긍정 신호 많음 |
| 3 | 프레피커피랩 | 0.8188 | 감점 없음 |
| 4 | 파티세리키세츠 | 0.8086 | 일반 불만 일부 반영 |
| 5 | 후니네김밥 | 0.7733 | 긍정 신호 많음 |
| 6 | 예향정 인천송도점 | 0.6019 | 약한 불만 있지만 긍정도 많음 |
| 7 | 스시사쿠 송도점 | 0.5399 | 리뷰 수 많고 긍정 많으나 일부 불만 감점 |

텍스트 때문에 내려간 주요 식당:

| 식당 | 이유 |
|---|---|
| 애월몽 | 알바 의심, 리뷰조작 직접 언급 |
| 고집132 본점 | 리뷰 조작 직접 언급 |
| 정성순대 송도점 | 리뷰 이벤트 때문에 평점 믿지 말라는 취지 |
| 뽁식당 송도점 | 리뷰 이벤트로 점수 뻥튀기 언급 |
| 후라토식당 송도트리플스트리트점 | 광고 언급과 강한 불만 다수 |
| 김밥천국 AT센터점 | 불친절, 위생, 맛없음 등 강한 불만 |
| 스시로 송도점 | 최악, 바퀴벌레, 다신 안 감 등 강한 불만 |

---

# Part F. 프로젝트 이동

## 29. OneDrive에서 로컬 폴더로 이동

문제:

- OneDrive 아래에서 Playwright가 TXT/JSON/XLSX를 계속 저장하면 동기화 지연이 생길 수 있다.
- XLSX 저장 중 파일 잠금이나 느림이 발생할 수 있다.
- 대량 파일 생성 시 OneDrive 상태 아이콘과 동기화가 방해가 될 수 있다.

새 위치:

```text
C:\Users\kimwo\project_review
```

복사 명령:

```powershell
robocopy "C:\Users\kimwo\OneDrive\Desktop\대학 교재\반품향" "C:\Users\kimwo\project_review" /E /XD ".git" /R:2 /W:1
```

복사 후 이동:

```powershell
cd "C:\Users\kimwo\project_review"
```

테스트 명령:

```powershell
python xlsx_build\kakao_review_text_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_리뷰내용반영_이사후테스트.html"
```

---

# Part G. 다시 실행할 때 필요한 명령어 모음

## 30. 음식점 목록 다시 수집

```powershell
cd "C:\Users\kimwo\project_review"
$env:KAKAO_REST_API_KEY="REST_API_KEY_값"
python xlsx_build\kakao_restaurant_search.py "인천 연수구 송도과학로"
```

## 31. 리뷰 통합 XLSX 다시 수집

```powershell
node xlsx_build\kakao_place_visible_text.mjs --xlsx-file "outputs\kakao_restaurants\kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx" --scrolls 3 --fast-threshold 40 --max-scrolls 200 --target-coverage 0.95 --stall-limit 8 --stall-wait-ms 5000 --wait-ms 2500 --scroll-delay-ms 1200 --combined-xlsx "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx"
```

## 32. 3개 문제 식당만 재수집

```powershell
node xlsx_build\kakao_place_visible_text.mjs --url-file "outputs\kakao_place_text\retry_failed_3_urls.txt" --scrolls 3 --fast-threshold 0 --max-scrolls 300 --target-coverage 1.0 --stall-limit 15 --stall-wait-ms 10000 --wait-ms 3500 --scroll-delay-ms 1800 --combined-xlsx "outputs\kakao_place_text\재수집_실패3개.xlsx"
```

## 33. 숫자 기반 순위표 생성

```powershell
python xlsx_build\kakao_review_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_자동스크롤_현재파일.html"
```

## 34. 리뷰 내용 반영 순위표 생성

```powershell
python xlsx_build\kakao_review_text_ranker.py "outputs\kakao_place_text\송도과학로_통합리뷰_자동스크롤.xlsx" --out "outputs\kakao_place_text\송도과학로_순위표_리뷰내용반영.html"
```

## 35. 특정 카카오 place URL만 확인

```powershell
node xlsx_build\kakao_place_visible_text.mjs "https://place.map.kakao.com/1922201261#review" --scrolls 0 --no-auto-scroll --wait-ms 2500 --combined-xlsx "outputs\kakao_place_text\_place_1922201261_lookup.xlsx"
```

---

# Part H. 현재 최종 산출물

## 36. 카카오 음식점 목록

```text
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도과학로_20260623_133939.xlsx
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도과학로_20260623_133939.csv
outputs/kakao_restaurants/kakao_restaurants_인천_연수구_송도과학로_20260623_133939.json
```

## 37. 통합 리뷰

```text
outputs/kakao_place_text/송도과학로_통합리뷰_자동스크롤.xlsx
```

## 38. 숫자 기반 순위표

```text
outputs/kakao_place_text/송도과학로_순위표_자동스크롤_현재파일.html
outputs/kakao_place_text/송도과학로_순위표_자동스크롤_현재파일.csv
```

## 39. 리뷰 내용 반영 순위표

```text
outputs/kakao_place_text/송도과학로_순위표_리뷰내용반영.html
outputs/kakao_place_text/송도과학로_순위표_리뷰내용반영.csv
```

## 40. 문서

```text
docs/kakao_restaurant_analysis_timeline.md
docs/chat_full_work_timeline.md
```

---

# Part I. 주의점과 한계

## 41. 카카오 리뷰 수집의 한계

- 카카오가 모든 리뷰를 DOM에 붙여주지 않으면 목표치까지 도달하지 못할 수 있다.
- 일부 페이지는 로딩이 느려서 오래 기다려야 한다.
- `후기미제공`은 수집 실패가 아니라 카카오가 후기를 제공하지 않는 상태다.
- 검색어가 `송도과학로` 하나면 주변 식당이 누락될 수 있다.

## 42. 순위 모델의 한계

- 조작 의심 점수는 확정 판정이 아니라 위험 신호다.
- 별 5개가 몰렸다고 무조건 조작은 아니다.
- 이벤트, 신규 오픈, 방송 노출, 단체 방문도 리뷰 몰림을 만들 수 있다.
- 텍스트 감점은 키워드 기반이라 문맥을 완벽히 이해하지는 못한다.
- 그래서 HTML 표에서 대표 문구를 같이 보고 사람이 최종 판단하는 구조가 좋다.

## 43. 앞으로 개선하면 좋은 것

1. 카카오 Local API 검색어를 여러 개 돌려 누락 식당 줄이기
2. 같은 장소 ID 기준으로 여러 검색 결과 병합하기
3. 리뷰 텍스트 임베딩이나 LLM 기반 감성/조작 의심 분류 추가하기
4. 리뷰 날짜 히스토그램을 HTML에 시각화하기
5. 상위권 식당만 따로 비교하는 요약 HTML 만들기
6. 음식 카테고리별로 점수를 정규화하기
7. `덕이네`처럼 도로명 검색에 안 걸린 식당을 수동 후보로 추가하는 기능 만들기

---

## 44. 한 줄 결론

이 프로젝트는 단순히 카카오맵 별점 높은 식당을 찾는 도구가 아니라, 리뷰어 성향, 리뷰 수 신뢰도, 별 5개 리뷰의 날짜 패턴, 실제 리뷰 내용까지 함께 읽어서 “정말 괜찮을 가능성이 높은 식당”을 고르는 분석 도구로 발전했다.
