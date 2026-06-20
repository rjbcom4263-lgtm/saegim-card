# saegim-card 수정 이력

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
