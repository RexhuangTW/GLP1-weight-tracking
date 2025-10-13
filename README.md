# GLP1 Weight Tracking 

用一份 **數據源檔案**（CSV 或 Excel） + 一支 **Python 產生器**，自動輸出每週的：

- `weekly/<YYYY-CWNN>_weight_tracking.xlsx`
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weekly_report.md`
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weight_trend.png`
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_bodyfat_trend.png`

> 週期以 **週四為每週起始**，第一週的 **anchor** 預設為 `2025-08-15`（可改）。

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
- 將一天中的多次測量分類為「早上」（5:00-12:00）和「晚上」（其他時間）
- 若同一時段有多次測量，會自動計算平均值

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

---

## 📤 產出說明

- `weekly/<YYYY-CWNN>_weight_tracking.xlsx`：該週 Excel 快照（只含那週資料）。  
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weekly_report.md`：Markdown 週報（內含資料表、統計、建議及圖表引用）。  
- `reports/<YYYY-CWNN>/<YYYY-CWNN>_weight_trend.png`、`reports/<YYYY-CWNN>/<YYYY-CWNN>_bodyfat_trend.png`：該週趨勢圖。  
- `reports/summary/overall_summary_report.md`：總結報告（使用 `--summary` 參數時產生）。
- `reports/summary/summary_weight_trend.png`、`reports/summary/summary_bodyfat_trend.png`：總體趨勢圖。

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
- **晚上**：12:00 PM - 5:00 AM（隔天）
- 若同一時段有多次測量，會自動計算平均值

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

---

## 📝 小提示
- **CSV 格式**：直接從體脂計匯出，腳本會自動分類早上/晚上測量值。
- **Excel 格式**：需要手動維護早上/晚上的數據。
- 使用 `--summary` 參數可以產生總結報告，查看整體減重進度。
- 如果想把 anchor 改成其他日期（例如療程第二階段），只要改 `--anchor-date` 即可。
- 建議定期備份 CSV 檔案，避免數據遺失。

---

## 📄 授權
（可自行選擇 License；若未指定，建議 MIT）
