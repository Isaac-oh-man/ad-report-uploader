# GFA 보고서 -> 구글 시트 자동 업로드

## 구조
```
naver-ads-report/
├── config.py           # 시트 ID, 컬럼 매핑, 폴더 경로
├── csv_processor.py    # ZIP 해제 + CSV 로드
├── sheet_uploader.py   # 구글 시트 업로드
├── main.py             # 실행 (단발 / 감시 모드)
├── inbox/              # 여기에 CSV/ZIP을 넣으면 자동 처리
├── done/               # 처리 완료 파일 보관
├── credentials/        # Google 서비스 계정 JSON
├── .env.example
└── requirements.txt
```

## 대상 시트
- (TW 360) Ads Report_노랑통닭
- https://docs.google.com/spreadsheets/d/1MHMHKbdPKtf6RijQdWjPzTxA-z72XHDd_jxJ6xyUGdk

## 사전 준비

### 1. Google Sheets API 서비스 계정
1. Google Cloud Console에서 프로젝트 생성
2. Google Sheets API 활성화
3. 서비스 계정 생성 -> JSON 키 다운로드
4. credentials/service_account.json으로 저장
5. 스프레드시트에서 서비스 계정 이메일을 편집자로 공유

### 2. 환경 설정
```bash
cp .env.example .env
pip install -r requirements.txt
```

## 사용법

### 단발 실행 (inbox 폴더의 파일 처리)
```bash
python main.py
```

### 감시 모드 (파일 넣으면 자동 처리)
```bash
python main.py --watch
```

### 덮어쓰기 모드
```bash
python main.py --mode overwrite
```

## 워크플로우
1. GFA에서 소재별 보고서 다운로드 (ZIP/CSV)
2. inbox/ 폴더에 파일 넣기
3. 자동으로 구글 시트에 업로드
4. 처리된 파일은 done/ 폴더로 이동
