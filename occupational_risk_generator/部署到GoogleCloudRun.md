# 部署到 Google Cloud Run

這份專案已可部署到 Google Cloud Run。

## 需要準備

- Google 帳號
- Google Cloud 專案
- 已啟用 Billing
- 已啟用 Cloud Run 與 Cloud Build

## 建議部署方式：Cloud Shell

1. 到 Google Cloud Console：

```text
https://console.cloud.google.com
```

2. 開啟右上角的 Cloud Shell。
3. 上傳本專案 zip 檔，或將專案檔案放到 Cloud Shell。
4. 進入專案資料夾。
5. 執行：

```bash
gcloud run deploy occupational-risk-generator \
  --source . \
  --region asia-east1 \
  --allow-unauthenticated
```

`asia-east1` 是台灣附近的 Google Cloud 區域。

## 若不想公開給所有人

不要加 `--allow-unauthenticated`。

部署後可到 Cloud Run 的「權限」設定，指定可使用的 Google 帳號。

## 啟動檔

Cloud Run 會使用 `Dockerfile` 啟動 Streamlit：

```bash
streamlit run app.py --server.address=0.0.0.0 --server.port=${PORT}
```

## 注意

本工具可能輸入健康管理分級、健檢異常與職務內容。若包含個資或健康資料，不建議開放公開存取，建議限制登入帳號。
