# KNO Workbench (운영/prod) — 개발/배포 메모

## 구성
- **백엔드**: `Code.gs` + `appsscript.json` — Apps Script 프로젝트(스프레드시트 바인딩). `clasp`로 배포.
  - scriptId: `1rDmkJsilCwPxA-aCWHWlp2QT5oish-Eav7JiUT1KEeBGTkgb9L-Gj_EY`
  - 바인딩 스프레드시트: `KNO Workbench v1.0` (id `1Eq_E4LDAbQqGKnXwQ2S4Enq9MTDX1YwBdQcZjCATILE`) — 실제 연구원 데이터 있음
  - 웹앱 API 엔드포인트(`/exec`, 배포 ID `AKfycbxGQ25Q...`): 프론트가 `fetch`로 호출
- **프론트**: `index.html` — GitHub Pages(`https://korneo.github.io/KNO-workbench/`)에서 서빙. clasp 대상 아님(`.claspignore`가 `appsscript.json`·`Code.gs`만 push).
- 계정: git/gh = `KorNEO`, clasp(Apps Script/Drive) = `koreanneology@gmail.com`

## dev와의 관계
- `KNO-workbench-dev`는 신어 집필 부분만 떼어낸 **축소판**이며 스프레드시트에 `연구원` 시트가 없음 → 계정/로그인/비번 기능은 dev에서 테스트 불가.
- `Code.gs`는 dev와 prod가 **URL 상수 2개(WEBAPP_URL, PAGES_URL)만 다르고 나머지 동일**.
- 계정류 작업은 사실상 prod에서 직접 확인. (제대로 된 dev를 원하면 prod 스프레드시트 완전 복제본으로 재구성 필요.)

## 배포 절차
1. 백엔드: `clasp push --force` → `clasp deploy -i AKfycbxGQ25QDvzAdXOCdYWXihv3Lkdj6zVXyq5M0KiGjccGJTRbiY1XRMvRjCHKrmlFdWLZ -d "설명"`
   - 반드시 **기존 배포 ID(`-i`)** 에 재배포해야 `/exec` URL이 유지됨. 새로 만들면 URL이 바뀌어 프론트가 깨짐.
2. 프론트: `index.html` 커밋 → `git push` (GitHub Pages 자동 반영, 1~2분)

## ⚠️ 중요 함정 — 웹앱 익명 접근 리셋
- `appsscript.json`에 아래 **`webapp` 블록이 반드시 있어야** 함. 없으면 `clasp deploy` 시 접근 권한이 기본값(로그인 필요)으로 **리셋**되어, 프론트의 익명 `fetch`가 **`Failed to fetch`**(CORS/권한) 로 실패함.
  ```json
  "webapp": { "executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS" }
  ```
  - `executeAs: USER_DEPLOYING` = 소유자(koreanneology) 권한으로 실행 → 비공개 스프레드시트 접근 가능
  - `access: ANYONE_ANONYMOUS` = 로그인 없이 누구나 접근(익명)
- **재배포 후에도 접근 권한을 편집기에서 수동으로 다시 줘야 할 수 있음**(과거에도 발생). Apps Script 편집기 → 배포 → 배포 관리 → 액세스 권한 = "모든 사용자" 확인/재설정.

## 데이터 안전 원칙
- 비밀번호 관련 작업(재설정/관리자 초기화)은 **연구원 시트의 비밀번호 해시 컬럼(10번)만** 수정. 아이디·토큰·소속·성별·작업 데이터는 절대 건드리지 않음.
- 큰 변경 전 스프레드시트 사본 백업 권장. (예: `KNO Workbench v1.0 — 백업 YYYYMMDD`)

## 비밀번호 재설정 기능 (2026-07 추가)
- **셀프 재설정**: 로그인 화면 "비밀번호를 잊으셨나요?" → 이메일 OTP 인증 → 새 비번 설정. `resetPassword(email, code, pw)`.
- **관리자 수동 초기화**: 연구원 명부 표의 `계정` 열 "🔑 비번 초기화" → 임시 비번 발급. `adminResetPassword(token, email)` (매니저 전용).
- OTP 보안: 5회 오입력 시 무효화, 60초 재발송 쿨다운.

## 의미부 선별 집필 (2026-08-30, 백엔드 @33)
- 회의 결과: 형태부는 전 항목, 의미부(뜻풀이·용례)는 NN+YT 합산 단어빈도 10 이상 항목만 집필.
- 항목 시트 컬럼 `의미부 대상`: `Y`/`N`. **빈 값은 Y로 간주**(이전 프로젝트 호환). `N`이면 의미부 단계에서도 형태부만 점검·상태·집계(`semTarget_`), 프론트는 의미부 폼 잠금·`의미부 미집필(빈도 10미만)` 표시.
- 업로드 CSV는 `workbench/_tools/build_write_upload.py`(계열 배분 xlsx → 32컬럼, `--llm-csv`로 LLM 컬럼 채움). 배분은 `workbench_reference/_tools/series/distribute_series.py --sem-col '단어빈도 합계' --sem-min 10`.
