# Cloudflare D1 코드 적용 안내

## 현재 완료된 준비
- Cloudflare D1 `ablab-sales-db` 생성
- D1 테이블 생성
- Worker `ablab-sales-api` 배포
- Vercel 환경변수 `D1_WORKER_URL`, `D1_API_KEY` 등록

## 적용 순서
1. 현재 `sales-dashboard` 폴더를 한 번 더 백업합니다.
2. 이 ZIP의 파일과 폴더를 현재 `sales-dashboard` 안에 덮어씁니다.
3. VS Code 터미널을 `sales-dashboard` 경로에서 열고 실행합니다.

```powershell
npm install
npm run build
```

4. Worker 코드가 갱신되었으므로 이어서 실행합니다.

```powershell
cd cloudflare-worker
npm install
npm run deploy
cd ..
```

기존에 등록한 Worker Secret은 유지되므로 `npm run secret`을 다시 할 필요가 없습니다.

5. GitHub/Vercel 배포:

```powershell
git add .
git commit -m "Migrate Sales Report storage to Cloudflare D1"
git push
```

6. Vercel의 최신 배포가 Ready가 된 뒤 사이트에서 `Ctrl + F5`를 누릅니다.

## 첫 테스트
1. 관리자 로그인
2. 당월 매출 파일 1개 업로드
3. 저장 완료 후 새로고침
4. 같은 금액 유지 확인
5. 시크릿 모드에서 확인
6. Cloudflare D1 Studio의 `sales_records`, `sales_upload_batches`에 행이 생겼는지 확인
7. 이후 전월, 전년동월 순서로 한 번씩 업로드

## 중요
- 기존 Supabase 환경변수와 데이터는 테스트 완료 전까지 삭제하지 않습니다.
- D1 코드는 당월 매출을 날짜별로 누적하고, 같은 날짜를 다시 올릴 때만 해당 날짜를 교체합니다.
- 전월과 전년동월은 기준월 단위로 전체 교체합니다.
