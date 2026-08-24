# proposal-calculator-backend

建議書系統的後端試算服務。用 LibreOffice headless + PyUNO 直接重新計算富邦官方 Excel 建議書範本，
確保數字跟官方試算軟體 100% 一致（不是重寫公式，是真的載入原始 xlsx 讓 LibreOffice 算）。

## API

POST /api/{PFA3|PFA6}/calculate

Request JSON:
```json
{
  "name": "客戶姓名",
  "gender": "男或女",
  "birth_year": 110,      // 民國年
  "birth_month": 1,
  "birth_day": 1,
  "payment_term": 6,       // 3 或 6
  "payment_freq": "年繳",  // 年繳/半年繳/季繳/月繳
  "dividend_scenario": "中分紅", // 高分紅/中分紅/低分紅（僅影響PDF標示）
  "input_mode": "face_amount",  // 或 "premium"
  "face_amount_wan": 30,        // input_mode=face_amount 時必填，保額(萬)
  "premium_amount": 100000,     // input_mode=premium 時必填，年繳保費(元)
  "death_benefit_pct": 0,       // 選填，指定保險金分期比例(0~100)
  "installment_period": 20,     // 選填，10 或 20
  "relationship": "同被保險人", // 選填
  "discount": "無",             // 選填
  "want_pdf": true              // 選填，是否連同產出 PDF (base64)
}
```

Response JSON 包含：`premiums`（各期保費）、`tables`（三種法定分紅情境的完整年度表：
zero_possible=假設分紅金額可能為零、most_likely=最可能紅利、lower=較低紅利）、
以及 `pdf_base64`（若 want_pdf=true）。

## 部署（Railway）

1. Railway → New Project → Deploy from GitHub repo → 選這個 repo
2. Railway 會自動偵測 Dockerfile 並建置（安裝 LibreOffice 需要幾分鐘，正常）
3. 部署完成後在 Settings 產生 Public Domain，記下網址供前端呼叫
