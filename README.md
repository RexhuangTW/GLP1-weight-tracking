# GLP1 Weight Tracking 

用一份 **數據源檔案**（CSV 或 Excel） + 一支 **Python 產生器**，自動輸出每週的：

- `weekly/<YYYY-CWNN>_weight_tracking.xlsx`
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weekly_report.md`
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weight_trend.png`
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_bodyfat_trend.png`

> 以 `--anchor-date` 指定的日期作為每週的第一天；預設 anchor 為 `2025-08-15`（可改）。

---

## 📂 專案結構

```
GLP1-weight-tracking/
├─ BodyComposition_202507-202510.csv  # 體脂計匯出的 CSV 檔案（自動讀取）
├─ GLP1_weight_tracking_master.xlsx   # （可選）手動維護的 Excel 主檔
├─ generate_weekly_report.py          # 週報產生器（支援 CSV/Excel、中文顯示）
├─ weekly/                            # 產生的「該週 Excel」
└─ reports/                           # 產生的「週報 Markdown + 圖表」
    ├─ 2025-CW01/                     # 每週報告資料夾
    ├─ 2025-CW02/
    └─ summary/                       # 總結報告資料夾
```

---

## 💻 環境需求

- Python 3.8+
- 套件：`pandas`、`openpyxl`、`matplotlib`
- （Linux 推薦）中文字型：`fonts-noto-cjk`

### 安裝
```bash
python3 -m pip install --upgrade pip
python3 -m pip install pandas matplotlib openpyxl
# 中文字型（Ubuntu/Debian）
sudo apt-get update && sudo apt-get install -y fonts-noto-cjk
# 第一次安裝字型後，建議清除 matplotlib 快取
rm -rf ~/.cache/matplotlib
```

---

## 🗂️ 數據源格式

### 方式 1：使用 CSV 檔案（推薦）

直接從體脂計（如 OMRON HBF-222T）匯出 CSV 檔案，腳本會自動：
- 解析測量日期時間
- 將一天中的多次測量分類為「早上」（5:00-11:59）和「晚上」（12:00-隔天 4:59）
- 若同一時段有多次測量，會自動計算平均值
 - 凌晨 0:00–4:59 視為前一天的晚上（PM）
 - 自動計算衍生欄位：脂肪重量(kg)、骨骼肌重量(kg)

CSV 檔案需包含以下欄位：
- `測量日期`（格式：`YYYY/MM/DD HH:MM`）
- `體重(kg)`
- `體脂肪(%)`

### 方式 2：使用 Excel 檔案

在 Excel 檔案的 **Daily Log** 工作表中，手動維護以下欄位：
- `日期`
- `早上體重 (kg)`、`晚上體重 (kg)`
- `早上體脂 (%)`、`晚上體脂 (%)`
- （可選）`藥物劑量 (mg)`、`副作用紀錄`、`每日飲水量 (L)`

> 腳本內建常見別名對應（如 *AM weight / PM weight / 早上體重 / 晚上體重* 等）。

---

## ▶️ 使用方式

### 1) 使用 CSV 檔案（預設）- 產生最新一週報告
```bash
python3 generate_weekly_report.py BodyComposition_202507-202510.csv
```

### 2) 產生總結報告（從第一天到最新數據）
```bash
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary
```

### 3) 指定第 N 週（從 anchor 起算；1 = 2025-08-15～2025-08-21）
```bash
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index 1
```

### 4) 使用 Excel 檔案
```bash
python3 generate_weekly_report.py GLP1_weight_tracking_master.xlsx --sheet "Daily Log" --header-row 1
```

### 5) 自訂 anchor 日期和輸出目錄
```bash
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --anchor-date 2025-08-15 --out-root .
```

### 6) 產生月報（最新或指定月份）
```bash
# 產生最新月份月報
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --monthly

# 產生指定月份（YYYY-MM）
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --monthly 2025-09
```

### 7) 長期目標與 ETA（預估達標日期）
```bash
# 於週報 / 月報 / 總結加入「體重 79kg、體脂 12%」的長期目標
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary \
  --goal-weight 79 --goal-fat-pct 12

# 指定 ETA 算法（預設：--eta-scope global, --eta-metric fatkg, --eta-method endpoint_all）
# 例如改回「近 28 天回歸」（regress28）
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary \
  --goal-weight 79 --goal-fat-pct 12 --eta-method regress28
```

---

## ⚙️ 參數說明

| 參數 | 說明 | 範例 |
|---|---|---|
| `master` | 數據源檔案路徑（CSV 或 Excel） | `BodyComposition_202507-202510.csv` |
| `--sheet` | Excel 工作表名稱（僅用於 Excel） | `"Daily Log"` |
| `--header-row` | 標題列索引（僅用於 Excel，0=第一列） | `1` |
| `--anchor-date` | 每週起始的對齊基準日（**週四**），第一週從這天開始 | `2025-08-15` |
| `--week-index` | 第幾週（1-based；不給則抓最後一週） | `2` |
| `--out-root` | 輸出根目錄 | `.` |
| `--summary` | 產生總結報告（從第一天到最新數據） | （flag 參數，無需值） |
| `--monthly [YYYY-MM]` | 產生月報；不帶值則輸出最新月份 | `--monthly 2025-09` |
| `--goal-weight` | 長期目標體重 (kg) | `79` |
| `--goal-fat-pct` | 長期目標體脂率 (%) | `12` |
| `--eta-scope` | ETA 視窗：`global` 使用全資料最後日回推；`local` 僅用目前報告區間 | `global` |
| `--eta-metric` | ETA 指標：`fatkg`（脂肪重量, 預設）/ `weight` / `fatpct` | `fatkg` |
| `--eta-method` | ETA 方法：`regress28`（近28天回歸, 預設）/ `endpoint_all`（首末端點, 全期間）/ `regress_all`（全期間回歸）/ `endpoint28`（近28天端點） | `endpoint_all` |
| `--show-target-lines` | 在圖表上繪製目標參考線（預設不顯示） | （flag） |
| `--no-target-lines` | 不繪製目標參考線（預設） | （flag） |

---

## 📤 產出說明

- `weekly/<YYYY-CWNN>_weight_tracking.xlsx`：該週 Excel 快照（只含那週資料）。  
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weekly_report.md`：Markdown 週報（內含資料表、統計、建議及圖表引用）。  
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weight_trend.png`、`reports/<YYYY-CWNN>/<YYYY-CWNN>_bodyfat_trend.png`：該週趨勢圖。  
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_visceral_fat_trend.png`、`reports/<YYYY-CWNN>/<YYYY-CWNN>_muscle_trend.png`：內臟脂肪與骨骼肌趨勢圖。  
- `reports/monthly/<YYYY-MM>/<YYYY-MM>_monthly_report.md`：月報（含 KPI、分析與目標/ETA）。  
- `reports/summary/overall_summary_report.md`：總結報告（使用 `--summary` 參數時產生）。
- `reports/summary/summary_weight_trend.png`、`reports/summary/summary_bodyfat_trend.png`、`reports/summary/summary_visceral_fat_trend.png`、`reports/summary/summary_muscle_trend.png`：總體趨勢圖。

> 週碼 `YYYY-CWNN` 的年份取該週 **起始日** 年份；`NN` = `week-index`。

---

## 🧩 常見問題（FAQ）

### 1) `ModuleNotFoundError: No module named 'pandas'`
請先安裝依賴：
```bash
python3 -m pip install pandas matplotlib openpyxl
```

### 2) 開啟圖表時中文顯示亂碼
安裝中文字型並清快取：
```bash
sudo apt-get install -y fonts-noto-cjk
rm -rf ~/.cache/matplotlib
```

### 3) CSV 檔案中的測量時間如何分類？
- **早上**：5:00 AM - 12:00 PM
- **晚上**：12:00 PM - 4:59 AM（隔天）
- 若同一時段有多次測量，會自動計算平均值
 - 凌晨 0:00–4:59 視為前一天的晚上（PM）

### 4) `⚠️ 無法從 Excel 映射必要欄位`（僅 Excel 格式）
代表程式抓錯標題列或工作表：
- 確認標題列是第幾列（通常是第 2 列 → `--header-row 1`）  
- 指定工作表名稱：`--sheet "Daily Log"`  
- 若欄位名稱不同，程式會嘗試別名對應；若仍失敗，請回報錯誤訊息中「偵測到的欄位」。

### 5) 如何產生所有週的報告？
使用迴圈產生所有週的報告：
```bash
for i in {1..9}; do
  python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index $i
done
```

### 6) 如何讓週/月/總結的 ETA 一致？
預設（`--eta-method endpoint_all`）為「第一筆到最新一筆」端點法。若希望改回「近 28 天線性回歸」，請加上 `--eta-method regress28`：
```bash
# 週報（第 1～9 週），示範改回 regress28
for i in {1..9}; do
  python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index $i \
    --eta-method regress28
done

# 月報（最新與指定月份）
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --monthly --eta-method regress28
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --monthly 2025-08 --eta-method regress28
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --monthly 2025-09 --eta-method regress28
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --monthly 2025-10 --eta-method regress28

# 總結
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary --eta-method regress28
```

---

## 📝 小提示
- **CSV 格式**：直接從體脂計匯出，腳本會自動分類早上/晚上測量值。
- **Excel 格式**：需要手動維護早上/晚上的數據。
- 使用 `--summary` 參數可以產生總結報告，查看整體減重進度。
 - 目標/ETA：預設長期目標為「體重 79kg、體脂 12%」，ETA 預設為 endpoint_all；可用 `--goal-weight`、`--goal-fat-pct` 與 `--eta-*` 覆寫。
- 如果想把 anchor 改成其他日期（例如療程第二階段），只要改 `--anchor-date` 即可。
- 建議定期備份 CSV 檔案，避免數據遺失。
 - 圖表包含「7 日移動平均」。目標線預設關閉，需顯示可加入 `--show-target-lines`。

---

## 📄 授權
（可自行選擇 License；若未指定，建議 MIT）
