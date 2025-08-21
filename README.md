# GLP1 Weight Tracking 

用一份 **主檔 Excel**（你手動更新） + 一支 **Python 產生器**，自動輸出每週的：

- `weekly/<YYYY-CWNN>_weight_tracking.xlsx`
- `reports/<YYYY-CWNN>_weekly_report.md`
- `reports/<YYYY-CWNN>_weight_trend.png`
- `reports/<YYYY-CWNN>_bodyfat_trend.png`

> 週期以 **週五為每週起始**，第一週的 **anchor** 預設為 `2025-08-15`（可改）。

---

## 📂 專案結構

```
GLP1-weight-tracking/
├─ GLP1_weight_tracking_master.xlsx   # 你手動維護的主檔（Daily Log）
├─ generate_weekly_report.py          # 週報產生器（支援中文顯示、週五起始）
├─ weekly/                            # 產生的「該週 Excel」
└─ reports/                           # 產生的「週報 Markdown + 圖表」
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

## 🗂️ 主檔（Daily Log）欄位

請在 `GLP1_weight_tracking_master.xlsx` 的 **Daily Log** 工作表中，維護以下欄位：

- `日期`
- `早上體重 (kg)`、`晚上體重 (kg)`
- `早上體脂 (%)`、`晚上體脂 (%)`
- （可選）`藥物劑量 (mg)`、`副作用紀錄`、`每日飲水量 (L)`

> 腳本內建常見別名對應（如 *AM weight / PM weight / 早上體重 / 晚上體重* 等），若仍對不上，請用 `--sheet` / `--header-row` 指定工作表與標題列。

---

## ▶️ 使用方式

### 1) 以 `2025-08-15（五）` 為第一週（anchor），抓「最後一週」
```bash
python3 generate_weekly_report.py GLP1_weight_tracking_master.xlsx --sheet "Daily Log" --header-row 1 --anchor-date 2025-08-15
```

### 2) 指定第 N 週（從 anchor 起算；1 = 2025-08-15～2025-08-21）
```bash
python3 generate_weekly_report.py GLP1_weight_tracking_master.xlsx --sheet "Daily Log" --header-row 1 --anchor-date 2025-08-15 --week-index 1
```

### 3) 指定輸出根目錄（預設為當前路徑）
```bash
python3 generate_weekly_report.py GLP1_weight_tracking_master.xlsx --out-root .
```

---

## ⚙️ 參數說明

| 參數 | 說明 | 範例 |
|---|---|---|
| `master` | 主檔路徑（可做為位置參數） | `GLP1_weight_tracking_master.xlsx` |
| `--sheet` | 主檔工作表名稱 | `"Daily Log"` |
| `--header-row` | 標題列索引（0=第一列） | `1` |
| `--anchor-date` | 每週起始的對齊基準日（**週五**），第一週從這天開始 | `2025-08-15` |
| `--week-index` | 第幾週（1-based；不給則抓最後一週） | `2` |
| `--out-root` | 輸出根目錄 | `.` |

---

## 📤 產出說明

- `weekly/<YYYY-CWNN>_weight_tracking.xlsx`：該週 Excel 快照（只含那週資料）。  
- `reports/<YYYY-CWNN>_weekly_report.md`：Markdown 週報（內含資料表、統計、建議及圖表引用）。  
- `reports/<YYYY-CWNN>_weight_trend.png`、`reports/<YYYY-CWNN>_bodyfat_trend.png`：趨勢圖。  

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

### 3) `⚠️ 無法從 Excel 映射必要欄位`
代表程式抓錯標題列或工作表：
- 確認標題列是第幾列（通常是第 2 列 → `--header-row 1`）  
- 指定工作表名稱：`--sheet "Daily Log"`  
- 若欄位名稱不同，程式會嘗試別名對應；若仍失敗，請回報錯誤訊息中「偵測到的欄位」，再補別名。

---

## 📝 小提示
- 你只要**持續手動更新** `GLP1_weight_tracking_master.xlsx`，其他週報與圖表都由腳本自動產生。  
- 如果想把 anchor 改成其他週五（例如療程第二階段），只要改 `--anchor-date` 即可。

---

## 📄 授權
（可自行選擇 License；若未指定，建議 MIT）
