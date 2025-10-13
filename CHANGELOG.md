# 更新日誌

## 2025-10-13：支援 CSV 格式

### 主要更新

#### 1. 新增 CSV 格式支援
- 腳本現在可以直接讀取體脂計（如 OMRON HBF-222T）匯出的 CSV 檔案
- 自動解析測量日期時間（格式：`YYYY/MM/DD HH:MM`）
- 自動分類測量時段：
  - **早上（AM）**：5:00 - 12:00
  - **晚上（PM）**：12:00 - 5:00（隔天）
- 若同一天同一時段有多次測量，自動計算平均值

#### 2. 保留 Excel 格式支援
- 原有的 Excel 讀取功能完全保留
- 可使用 `--sheet` 和 `--header-row` 參數指定工作表和標題列
- 欄位映射和別名對應功能維持不變

#### 3. 新增功能
- `--summary` 參數：產生從第一天到最新數據的總結報告
- 總結報告包含：
  - 總體統計數據
  - 完整時間範圍的趨勢圖
  - 平均每週體重變化
  - 重點成果分析

#### 4. 改進報告結構
- 週報現在儲存在 `reports/<週次>/` 子目錄中
- 總結報告儲存在 `reports/summary/` 目錄中
- 每個週次的所有檔案（MD、PNG）集中管理

### 技術細節

#### CSV 讀取邏輯
```python
# 讀取 CSV
df_raw = pd.read_csv(master_path)

# 解析時間
df_raw['測量日期時間'] = pd.to_datetime(df_raw['測量日期'], format='%Y/%m/%d %H:%M')
df_raw['小時'] = df_raw['測量日期時間'].dt.hour

# 分類時段
df_raw['時段'] = df_raw['小時'].apply(lambda h: 'AM' if 5 <= h < 12 else 'PM')

# 按日期和時段分組，計算平均值
```

#### 數據品質處理
- 自動處理缺失值（顯示為 `nan`）
- 統計計算時只使用有效數據
- 趨勢圖自動跳過缺失點

### 使用範例

#### 使用 CSV 檔案（推薦）
```bash
# 產生最新週報
python3 generate_weekly_report.py BodyComposition_202507-202510.csv

# 產生總結報告
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary

# 產生特定週次
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index 5
```

#### 使用 Excel 檔案（向後相容）
```bash
python3 generate_weekly_report.py GLP1_weight_tracking_master.xlsx --sheet "Daily Log" --header-row 1
```

### 測試結果

✅ 成功讀取 112 筆測量記錄（2025/08/15 - 2025/10/13）  
✅ 成功產生 9 週的週報（2025-CW01 到 2025-CW09）  
✅ 成功產生總結報告  
✅ 體重數據：98.7 - 109.6 kg，總減重 10.9 kg  
✅ 圖表正常顯示，中文字型無亂碼  

### 檔案清單

#### 新增/修改的檔案
- ✅ `generate_weekly_report.py` - 更新主腳本以支援 CSV
- ✅ `README.md` - 更新使用說明
- ✅ `USAGE.md` - 新增快速使用指南
- ✅ `CHANGELOG.md` - 本文件

#### 數據檔案
- `BodyComposition_202507-202510.csv` - 體脂計匯出的原始數據（112 筆記錄）

#### 產出檔案
- `weekly/2025-CW01_weight_tracking.xlsx` ~ `2025-CW09_weight_tracking.xlsx` - 各週 Excel 數據
- `reports/2025-CW01/` ~ `2025-CW09/` - 各週報告目錄（含 MD 和 PNG）
- `reports/summary/` - 總結報告目錄

### 向後相容性

✅ 完全向後相容  
- 原有的 Excel 讀取功能保持不變
- 所有參數和選項繼續有效
- 現有的工作流程無需修改

### 未來改進建議

1. **自動化檢測**：自動檢測檔案格式（CSV/Excel）並使用適當的讀取方法
2. **數據驗證**：新增數據完整性檢查和異常值偵測
3. **多格式匯出**：支援匯出 PDF 或 HTML 格式的報告
4. **互動式圖表**：使用 Plotly 產生可互動的圖表
5. **數據分析**：新增更多統計分析（趨勢線、相關性分析等）

---

## 舊版本記錄

### 初始版本
- 支援 Excel 格式讀取
- 週報產生功能
- 中文字型支援
- 週五為起始日的自訂週期
