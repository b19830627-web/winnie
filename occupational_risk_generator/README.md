# 特約職護職業安全風險評估產生器

這是一個可在本機執行的 Streamlit 網頁工具。輸入公司、產業、部門、職務、作業內容、班別、特殊健康風險與健康管理分級後，程式會依據 `risk_rules.json` 自動產生正式、專業的繁體中文建議文字，並可複製或下載為 Word 檔。

## 專案結構

```text
occupational_risk_generator/
├── app.py
├── risk_rules.json
├── requirements.txt
└── README.md
```

## 安裝方式

建議使用 Python 3.10 以上版本。

```bash
cd occupational_risk_generator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows PowerShell 啟用虛擬環境：

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 執行指令

```bash
streamlit run app.py
```

啟動後，瀏覽器會開啟本機網址，通常為：

```text
http://localhost:8501
```

## Mac 雙擊啟動

也可以直接雙擊：

```text
啟動職護風險評估產生器.command
```

若 macOS 顯示安全性提醒，請在檔案上按右鍵，選擇「打開」。程式啟動後會開啟本機網頁，網址通常是：

```text
http://localhost:8501
```

## 部署成線上版

本專案可部署到 Streamlit Community Cloud。部署時請將下列檔案放到同一個 GitHub Repository：

- `app.py`
- `risk_rules.json`
- `requirements.txt`
- `README.md`
- `.streamlit/config.toml`

部署時 Main file path 請填：

```text
app.py
```

詳細步驟請見：

```text
線上部署說明.md
```

## 使用方式

1. 選擇輸出格式：「四大段建議」或「附表八紀錄式」。
2. 填寫公司名稱、產業別、部門、職務名稱。
2. 輸入主要作業內容，例如久站、久坐、搬運、電腦作業、輪班、高溫、噪音、粉塵、化學品接觸等。
3. 選擇班別、特殊健康風險與健康管理分級。
4. 如需產生附表八紀錄，可展開「附表八紀錄欄位（選填）」填寫健康編號、健康評估依據、中高齡評估與適工性評估結果。
5. 如有現場特殊狀況，可填寫補充說明。
6. 按下「產生建議」。
7. 可直接複製完整文字，或下載 Word 檔。

## 附表八紀錄式輸出

選擇「附表八紀錄式」後，系統會產生下列結構：

- 基本資料與作業特性
- 健康評估與風險分析
- 改善建議與採行措施
- 管理建議
- 環境建議
- 教育指導
- 適工評估與後續追蹤

## 規則庫維護

風險規則集中於 `risk_rules.json`，每一筆規則包含：

- `match.fields`：要比對的表單欄位。
- `match.keywords`：觸發規則的關鍵字。
- `risk`：職業安全衛生風險評估文字。
- `education`：衛教建議。
- `improvements`：具體改善措施。
- `follow_up`：後續追蹤建議。

新增規則範例：

```json
{
  "id": "example_rule",
  "label": "範例風險",
  "match": {
    "fields": ["work_content"],
    "keywords": ["範例關鍵字"]
  },
  "risk": ["此處填寫風險評估文字。"],
  "education": ["此處填寫衛教建議。"],
  "improvements": ["此處填寫具體改善措施。"],
  "follow_up": ["此處填寫後續追蹤建議。"]
}
```

若輸入資料不足，程式會顯示「建議補充作業內容、班別或暴露因子，以利完整評估」，避免自行產生未被輸入資料支持的風險。
