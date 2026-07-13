# Sales Report V3 1단계 적용 안내

## 중요한 원칙
- 기존 `app_state` 테이블과 기존 데이터는 삭제하지 않습니다.
- 새 구조는 월별로 점진적으로 전환됩니다.
- 새 SQL이 적용되지 않은 상태에서도 코드는 기존 저장 방식으로 자동 복귀합니다.
- 새 구조에서 업로드가 실패하면 기존 매출 데이터는 바뀌지 않습니다.

## 파일 역할
- `supabase/sales_v3_phase1.sql`: Supabase SQL Editor에서 실행
- 나머지 프로젝트 파일: VS Code의 기존 프로젝트에 덮어쓰기

## 적용 순서

### 1. Supabase SQL 실행
1. Supabase에서 `ablabkorea-dailysales` 프로젝트를 엽니다.
2. 왼쪽 `SQL Editor`를 누릅니다.
3. `New query` 또는 빈 쿼리 창을 엽니다.
4. `supabase/sales_v3_phase1.sql`의 전체 내용을 복사해 붙여넣습니다.
5. `Run`을 누릅니다.
6. `Success`가 나오면 다음 단계로 이동합니다.
7. 빨간 오류가 나오면 코드 배포를 중지하고 오류 화면을 확인합니다.

### 2. 테이블 확인
`Table Editor`에서 아래 테이블이 보여야 합니다.
- `sales_records`
- `sales_upload_batches`
- 기존 `app_state`도 그대로 있어야 합니다.

### 3. 프로젝트 덮어쓰기
1. ZIP 압축을 풉니다.
2. 압축을 푼 폴더 안의 파일 전체를 복사합니다.
3. 현재 운영 프로젝트 `sales-dashboard` 폴더에 붙여넣습니다.
4. 파일 교체 질문에는 `대상 폴더의 파일 바꾸기`를 선택합니다.

### 4. VS Code 확인
터미널에서 아래 명령을 한 줄씩 실행합니다.

```powershell
npm install
```

```powershell
npm run build
```

빌드가 성공한 경우에만 다음으로 진행합니다.

### 5. Git 배포
```powershell
git add .
```

```powershell
git commit -m "Add transactional Sales V3 storage"
```

```powershell
git push
```

### 6. 사이트 확인
1. Vercel 배포가 `Ready`인지 확인합니다.
2. 사이트를 새로고침합니다.
3. 관리자 로그인 → `월초관리` → `업로드 관리`로 이동합니다.
4. 우측 상단에 `안전 저장 V3`가 표시되는지 확인합니다.

`기존 저장 방식`이 표시되면 SQL이 아직 적용되지 않았거나 Supabase가 503 상태입니다.

### 7. 점진적 데이터 전환
기존 데이터는 그대로 표시됩니다. 다시 올린 기간만 V3로 전환됩니다.

권장 순서:
1. 당월 파일 업로드
2. 새로고침 후 금액 유지 확인
3. 시크릿 모드 확인
4. 전월 파일 업로드
5. 다시 확인
6. 전년동월 파일 업로드
7. 다시 확인

## 새 구조의 안전장치
- 업로드는 Supabase 함수 안에서 한 번에 처리됩니다.
- 새 데이터 저장 중 오류가 나면 작업 전체가 취소됩니다.
- 기존 행은 삭제하지 않고 `active=false`로 보관합니다.
- 업로드 이력은 `sales_upload_batches`에 남습니다.
- 같은 날짜의 당월 업로드만 교체됩니다.
- 전월·전년동월은 같은 기준월 자료만 교체됩니다.
- 아직 V3로 전환하지 않은 기간은 기존 `app_state` 데이터가 계속 표시됩니다.

## 하지 말아야 할 작업
- `app_state` 삭제
- 기존 매출 키 삭제
- localStorage 초기화
- SQL 오류가 난 상태에서 전체 파일 재업로드
- 여러 기간 파일을 한꺼번에 연속 업로드
