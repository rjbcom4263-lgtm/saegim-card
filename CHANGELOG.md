# saegim-card 수정 이력

## 2026-06-21 — Firebase Firestore 이관 1단계 완료

### 배경

Apps Script + Google Sheet + Firebase 혼합 구조에서 반복적인 문제 발생:
- cd.html 무한 로딩
- QR 코드 중복 생성 (CD-0025)
- 고객용/관리자용 API 혼선
- QR_DB와 제품별 시트 동기화 오류

**결정**: Firestore를 메인 DB로 전환, Google Sheet는 백업/조회용으로 격하.

---

### 신규 구조

```
Firestore qr_cards/{code}          ← 공개 데이터 (고객 화면용)
Firestore qr_card_private/{code}   ← 민감 데이터 (비밀번호·이메일·푸시토큰)
Google Sheet QR_DB                 ← 백업·조회용
Firebase /admin                    ← 관리자 화면 (다음 단계)
```

---

### 작업 1: cd.html 파일 잘림 버그 수정

**원인**
`cd.html`이 286행에서 파일이 잘려 있었음. `val()` 함수 중간에서 끊겼고 `esc()`, `renderError()`, `loadData()` 호출이 모두 없었음. 브라우저가 스크립트 파싱 오류로 화면을 렌더링하지 못해 "불러오는 중입니다..."에서 무한 로딩 발생.

**수정 내용** (`cd.html` 말미에 추가)
```js
function val(id){const el=document.getElementById(id);return el?el.value.trim():'';}
function esc(s){const d=document.createElement('div');d.textContent=s==null?'':String(s);return d.innerHTML;}
function renderError(msg){app.innerHTML='<div class="error">'+esc(msg)+'</div>';}

loadData();
```

**규칙**: 파일 말미에 반드시 `loadData()` 호출이 있어야 화면이 뜸.

---

### 작업 2: admin_Code.gs QR 중복 방지 패치

**원인**
`getNextNumber_()` 함수가 제품별 시트만 보고 다음 번호를 계산 → QR_DB에 이미 있는 코드를 다시 생성하는 중복 발생 (CD-0025 2개 생성).

**수정 내용** (`admin_Code.gs`)

```js
// 제품 시트 + QR_DB 둘 다 스캔해서 max 계산
function getNextNumber_(sheet, productType) {
  const masterSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MASTER_SHEET_NAME);
  let max = 0;
  function scanSheet_(s) { /* ...코드별 최대 번호 추출... */ }
  scanSheet_(sheet);
  scanSheet_(masterSheet);
  return max + 1;
}
```

추가로 `codeExistsInMaster_()` 함수 + `createQrInventory()` while 루프로 생성 중 중복 건너뛰기 적용.

---

### 작업 3: QR_DB → Firestore 이관 (admin_Code.gs)

**추가 함수**

| 함수 | 역할 |
|---|---|
| `findDuplicateQrCodes()` | QR_DB 중복 코드 검사 |
| `migrateQrDbToFirestoreV1()` | qr_cards + qr_card_private 분리 이관 |
| `migrateOneQrRowV1_()` | 단일 행 분리 이관 |
| `writeFirestoreDoc_()` | Firestore REST API patch (토큰 재사용) |
| `toFirestoreFields_()` | 값 → Firestore fields 포맷 변환 |
| `pickFields_()` | 필드 목록 기반 객체 추출 |

**공개 컬렉션** `qr_cards` — 고객 화면에 필요한 모든 필드 (비밀번호 제외)

**비공개 컬렉션** `qr_card_private` — password, admin_password, owner_email, push_token, memo

**토큰 최적화**: 행마다 토큰을 재발급하지 않고 이관 시작 시 1회만 발급해 전 행에 재사용.

---

### 이관 결과

```
실행 함수: migrateQrDbToFirestoreV1
결과: 분리 이관 완료 — 성공 29건 / 스킵 0건 / 실패 0건

Firestore qr_cards        → 29개 문서 (CD, BMT, GT, WT, TT)
Firestore qr_card_private → 29개 문서 (password, admin_password, owner_email, push_token, memo)
```

---

### 다음 단계

```
1. Firestore 보안 규칙 설정
   - qr_cards: 읽기 공개 / 쓰기 인증 필요
   - qr_card_private: 읽기·쓰기 모두 인증 필요

2. Firebase /admin 관리자 화면 제작
   - public/admin/index.html
   - Firebase Auth (Google 로그인, 허용 이메일 제한)
   - Firestore 직접 읽기/쓰기

3. cd.html Firestore 직접 조회로 전환
   - Apps Script API 제거
   - Firebase SDK로 qr_cards/{code} 읽기

4. BMT → TT → WT → GT 순서로 전환
```

---

## 2026-06-20 — CD QR 등록 완료 기능 복구 및 전체 정비

### 배경

Firebase Hosting(`saegim-memory.web.app`)에 배포된 CD 제품 QR 카드에서
"등록 완료" 버튼을 눌러도 아무 반응이 없는 문제가 발생.

---

### 문제 1: git stash pop 충돌로 conflict marker가 HTML 파일에 삽입됨

**원인**  
이전 세션에서 `git stash pop`을 실행했는데, stash가 생성된 시점(commit `9b921e9`)과
remote에 새로 추가된 HTML 파일들이 충돌하면서 `<<<<<<<`, `=======`, `>>>>>>>` 마커가
`cd.html`, `bmt.html`, `gt.html`, `tt.html`, `index.html`에 삽입된 채로 커밋됨.

**해결**  
`git show a7bf51d:"<파일명>"` 명령으로 충돌 전 clean 버전을 복원 후 재커밋.

```bash
for f in cd.html bmt.html gt.html tt.html index.html; do
  git show a7bf51d:"$f" > "$f"
done
git add .
git commit -m "fix: conflict marker 제거, clean 버전 복원"
```

---

### 문제 2: API URL 오염 (관리자 URL이 고객용 파일에 적용됨)

**원인**  
Apps Script 배포 URL을 두 개 혼용 중:
- 고객용(customer-facing): `AKfycbyhOt-UJ_lZ8Upyw9ZvGKCNT0B2BEM3asnwOuI9_q7GsHHCYglWkPYmgPoz7FvNcNua`
- 관리자용(admin): `AKfycbyfXy...`

URL 교체 작업 중 관리자 URL이 고객용 HTML 파일 전체에 잘못 적용됨.
결과적으로 고객 페이지가 관리자 API를 호출하게 됨.

**해결**  
Chrome 내비게이션으로 두 URL의 응답을 직접 확인해 올바른 URL 식별 후 전체 파일에 재적용.

올바른 고객용 API URL:
```
https://script.google.com/macros/s/AKfycbyhOt-UJ_lZ8Upyw9ZvGKCNT0B2BEM3asnwOuI9_q7GsHHCYglWkPYmgPoz7FvNcNua/exec
```

---

### 문제 3: CD-0023 status가 '사용중'으로 등록 차단됨

**원인**  
`Code.gs`의 `registerSoldQr()` 함수는 `status === '판매완료'`인 경우에만 등록을 허용.
CD-0023의 status가 이미 `사용중`으로 되어 있어 등록 불가.

**해결**  
Google Sheets QR_DB 시트에서 CD-0023 행의 status를 `판매완료`로 수동 변경.

---

### 문제 4: `registerFirst()` 에러가 조용히 사라짐 + CORS preflight 차단

**원인**  
두 가지 복합 문제:

1. `apiPost()`가 `Content-Type: application/json`으로 POST 전송 시, 브라우저가
   CORS preflight OPTIONS 요청을 먼저 보내는데 Google Apps Script가 이를 처리하지 못해 fetch 실패.
2. `registerFirst()` 함수에 try-catch가 없어서 fetch 에러가 발생해도
   UI에 아무 반응이 없었음.

**해결** (`cd.html`)

`apiPost()` — Content-Type을 `text/plain`으로 변경해 preflight 우회:
```js
// 수정 전
async function apiPost(payload){
  const r=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  return r.json();
}

// 수정 후
async function apiPost(payload){
  const r=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify(payload),redirect:'follow'});
  return r.json();
}
```

`registerFirst()` — try-catch 추가:
```js
async function registerFirst(){
  try{
    // ... 기존 로직 ...
    const r=await apiPost(form);
    if(!r.success){alert(r.message||'등록 실패');return;}
    alert('등록이 완료되었습니다!');
    location.href='?code='+encodeURIComponent(CODE);
  }catch(e){alert('오류: '+e.message);}
}
```

---

### 결과

- CD-0023 QR 등록 완료 정상 작동 확인 ✅
- Firebase 배포 commit: `bde912a`

---

### admin_Code.gs 주요 업데이트 (별도 Apps Script 프로젝트)

> 아래 코드는 관리자 Apps Script 에디터에 직접 붙여넣어야 함.
> Firebase 배포 파일과 무관.

**추가된 상수**
```js
const ACTIVE_PRODUCT_TYPES = ['CD', 'BMT', 'GT', 'WT', 'GM'];
```

**확장된 QR_HEADERS** — 신규 제품(GM, WT, GT)용 컬럼 포함 총 60개 이상 헤더 정의.

**새 헬퍼 함수**
- `ensureSheetHeaders_()` — 기존 데이터 유지하면서 누락 헤더만 추가
- `makeQrRow_()` — 헤더 인덱스 기반으로 row 배열 생성 (컬럼 순서 변경에 안전)
- `initializeActiveSheets()` — 앱스 스크립트 에디터에서 1회 실행해 전체 시트 헤더 정리

---

### git 워크플로우 개선

**기존 문제**: VS Code가 `.git/index.lock`을 물고 있어 CMD push가 자주 실패.

**해결책**: GitHub Desktop 설치 및 연결 완료.

표준 push 순서 (CMD 방식 유지 시):
```bash
del C:\Users\새김\saegim-card\.git\index.lock
git add .
git commit -m "메시지"
git push origin main
```

GitHub Desktop 방식: Changes 탭 → Summary 입력 → Commit to main → Push origin

---

### Firebase 캐시 우회 팁

배포 후 브라우저가 이전 버전을 캐싱하는 경우 URL에 `?v=숫자` 파라미터를 붙여 확인:
```
https://saegim-memory.web.app/cd.html?code=CD-0023&v=3
```
