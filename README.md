# 새김 카드 (Saegim Card)

QR 기반 메모리 카드 서비스 — Firebase Hosting으로 운영되는 웹 앱입니다.

## 서비스 개요

반려동물·식물 등의 정보를 QR 코드로 새겨 태그를 만들고, 스캔하면 추억 페이지가 열리는 서비스입니다.

## 주요 기능

- **QR 태그 페이지** (`/wt.html`) — 바램 태그 스캔 후 표시되는 메모리 카드 페이지
- **어드민 패널** (`/admin/`) — QR 재고 관리, QR 생성, 스티커 출력, 사주 탭
- **안심뱃지 출력판** (`/admin/badge.html`) — 58파이 뱃지 도안 + QR 자동 배치 + 인쇄 출력
- **정원 빌더** (`/gt.html`) — 아이소메트릭 타일 기반 추모 정원
- **사주** (`/saju/`) — 사주팔자 분석 앱 (임베드)
- **개인정보 처리방침** (`/privacy.html`)

## 기술 스택

| 영역 | 기술 |
|------|------|
| 호스팅 | Firebase Hosting |
| 데이터베이스 | Cloud Firestore |
| 인증 | Firebase Auth (Google 로그인) |
| 스토리지 | Firebase Storage |
| 관리 API | Google Apps Script (`Code.gs`, `admin_Code.gs`) |
| 푸시 알림 | OneSignal |
| 프론트엔드 | 바닐라 HTML/JS, Canvas API |
| 사주 앱 | React + Vite (빌드 후 `/saju/`에 배포) |

## 프로젝트 구조

```
saegim-card/
├── index.html              # 루트 리다이렉트
├── wt.html                 # 바램 태그 메모리 카드
├── gt.html                 # 아이소메트릭 추모 정원
├── privacy.html            # 개인정보 처리방침
├── admin/
│   ├── index.html          # 어드민 패널
│   └── badge.html          # 안심뱃지 출력판
├── saju/                   # 사주팔자 앱 (빌드 결과물)
├── Code.gs                 # Google Apps Script (사용자용)
├── admin_Code.gs           # Google Apps Script (어드민용)
├── firebase.json           # Firebase 설정
└── storage.rules           # Firebase Storage 보안 규칙
```

## 로컬 개발

```bash
# Firebase CLI 설치
npm install -g firebase-tools

# 로그인
firebase login

# 로컬 서버 실행
firebase serve
```

## 배포

```bash
firebase deploy
```

## 어드민 접근

Firebase Auth Google 로그인 후 `ADMIN_EMAILS` 목록에 등록된 계정만 접근 가능합니다.  
Firestore 보안 규칙의 `isAdmin()` 함수에도 동일하게 등록해야 합니다.

## 사주 앱 빌드 & 배포

사주팔자 앱을 수정한 경우 아래 절차로 다시 빌드합니다.

```bash
cd "D:\사주 팔자"
npm run build
xcopy /E /I /Y "dist" "C:\Users\새김\saegim-card\saju"
```

이후 `firebase deploy`로 배포합니다.
