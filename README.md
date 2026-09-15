# 📦 포장 사양 · 견적 요청 대시보드

성우전자(본사) · 성우비나(베트남)의 제품 **포장 사양**을 입력하면
- 포장 방식별 **적재 효율**을 자동 계산하고
- 구매팀 전달용 **표준 견적 요청서**(Excel / PDF)를 생성하며
- 구매팀이 회신하는 **구매 확정 단가**를 입력·축적할 수 있는 Streamlit 대시보드입니다.

## 실행 방법

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 화면 구조

| 영역 | 내용 |
|------|------|
| 사이드바 | 제품 사이즈(L·W·H), 포장 방식(REEL/Bulk/Tray), 대상 법인 선택 + 법인별 **환율·인건비 가중치** 오버라이드 |
| 탭① 적재 효율 | 박스별 적재 수량·부피효율 표 + 막대그래프 |
| 탭② 견적 양식 | 구매팀 전달용 표준 양식 + **구매 확정 단가 입력란** + Excel/PDF 다운로드 |
| 탭③ 기준 데이터 | 적용 중인 박스 카탈로그 / 법인 기준값 열람 |

## 핵심 계산 공식 (적재 수량)

```
n_L = ⌊박스내경_L ÷ 제품외경_L⌋
n_W = ⌊박스내경_W ÷ 제품외경_W⌋
n_H = ⌊박스내경_H ÷ 제품외경_H⌋
박스당 적재수량 = n_L × n_W × n_H          (⌊ ⌋ = 내림)
```
- `data.py`에 정의된 **박스 내경(inner_l/w/h)** 이 기준입니다.
- `최적 방향(회전)` 옵션 ON 시 제품을 6가지 방향으로 놓아보고 최대 적재수량을 채택합니다.

## 원가 계산 (법인별 가중치 반영)

```
포장 인건비(가중) = 박스당 포장 인건비 × 인건비 가중치
박스당 총원가     = 박스 단가 + 포장 인건비(가중)
견적 통화 환산     = 박스당 총원가 × 환율(fx_rate)
```

## 구매 확정 단가 · 견적 이력 영구 저장 (Google Sheets 연동)

Streamlit Cloud는 파일 시스템이 임시(ephemeral)라서, 연동 없이는 앱을 재배포할 때마다
`records.csv`/`box_prices.csv`가 초기화됩니다. Google Sheets를 연결하면 재배포해도
데이터가 유지됩니다.

**설정 방법**
1. Google Cloud Console에서 서비스 계정을 만들고 JSON 키를 발급받습니다 (Sheets API 활성화).
2. Google Sheets에서 새 스프레드시트를 만들고, 위 서비스 계정 이메일을 **편집자**로 공유합니다.
3. Streamlit Cloud 앱 설정(**Settings → Secrets**)에 아래 형식으로 입력합니다.

   ```toml
   gsheet_id = "스프레드시트 URL의 /d/와 /edit 사이 ID"

   [gcp_service_account]
   type = "service_account"
   project_id = "..."
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "...@....iam.gserviceaccount.com"
   client_id = "..."
   token_uri = "https://oauth2.googleapis.com/token"
   ```

4. 저장 후 앱을 재실행하면 `records`/`box_prices` 시트가 자동 생성되며 이후 모든 저장이
   Google Sheets로 이루어집니다.

secrets가 설정되어 있지 않으면 기존처럼 로컬 CSV(`records.csv`/`box_prices.csv`)로 자동
폴백하므로, 로컬 개발 시에는 별도 설정 없이 그대로 실행할 수 있습니다.

## 데이터 관리 (수정 위치)

모든 기준 데이터는 **`data.py` 한 곳**에서 딕셔너리로 관리합니다.

- `ENTITIES` — 법인별 환율/인건비 가중치/포장 인건비
- `BOX_CATALOG[포장방식][법인코드]` — 표준 박스 리스트 (내경·단가·허용중량)

> 박스를 추가하려면 해당 리스트에 dict 한 줄만 추가하면 대시보드에 즉시 반영됩니다.
> 성우전자(`SW-KR`)와 성우비나(`SW-VN`) 데이터는 완전히 분리되어 있습니다.

## 파일 구성

```
app.py            Streamlit UI (사이드바 + 메인 대시보드)
data.py           기준 데이터 (박스 카탈로그, 법인 정보) ← 여기서 관리
calculations.py   적재 수량 / 부피효율 / 원가 계산 로직
exporters.py      Excel · PDF 견적서 생성 (한글 지원)
requirements.txt  의존성
```
