# Changelog

All notable changes to this project will be documented in this file.

## [2.0.0] - 2025-12-09

### 🎉 重大更新：全新醫療風格 Web 應用

完全重寫整個專案，從 Python 報告生成系統轉變為現代化的單頁 Web 應用程式。

### ✨ 新增功能

#### Phase 1 - 核心功能
- **CSV 資料匯出** - 一鍵匯出所有數據為 CSV 格式
- **表格筆數選擇器** - 可選擇顯示 15/30/50 筆或全部資料
- **深色模式** - 支援深色/淺色主題切換，自動儲存偏好設定
- **週比較分析** - 本週 vs 上週的體重、體脂、肌肉率、內臟脂肪對比

#### Phase 2 - 進階功能
- **進階目標設定** - 新增體脂率和 BMI 目標設定
- **最佳紀錄統計** - 顯示最低/最高體重、最低體脂率、最高肌肉率
- **波動分析** - 體重/體脂波動度（標準差）、最大單日變化、最長連降紀錄
- **響應式優化** - 大幅改善手機版顯示效果

#### Phase 3 - 高級功能
- **PDF 報告匯出** - 使用 jsPDF 生成完整的視覺化報告
- **數據輸入/編輯** - 本地新增和編輯體重紀錄功能
- **每日提醒** - 可設定每日推送通知提醒記錄體重
- **離線支援** - Service Worker 實現離線瀏覽
- **圖表互動增強** - 支援滾輪縮放、拖曳平移、雙擊重置

### 🎨 介面改進

- **醫療風格設計** - 採用專業醫療藍色配色方案（#2563EB, #0891B2, #059669）
- **醫療卡片組件** - 統一的 .medical-card 樣式
- **線條圖示** - 使用 Font Awesome Regular (far) 線條圖示
- **簡約動畫** - 移除過度裝飾，保留實用的淡入效果
- **InBody 人體圖** - 專業的身體成分視覺化

### 📊 資料視覺化

- **互動式圖表** - Chart.js v4 實現專業圖表
- **多維度分析** - 體重、體脂、肌肉率、內臟脂肪趨勢
- **7日移動平均** - 平滑的趨勢線顯示
- **目標線顯示** - 視覺化目標達成進度
- **即時篩選** - 7天/30天/90天/全期間快速切換

### 🔧 技術特點

- **單一檔案架構** - 完整功能封裝在單一 HTML 檔案中（2855 行）
- **CDN 依賴** - 所有套件透過 CDN 載入，無需本地安裝
- **LocalStorage 快取** - 30 分鐘資料快取減少網路請求
- **Google Sheets 整合** - 即時讀取線上試算表資料
- **漸進式 Web App** - 支援離線使用和通知推送

### 📱 響應式設計

- **手機優先** - 針對行動裝置優化的佈局
- **自適應表格** - 水平滾動處理大量欄位
- **彈性網格** - 根據螢幕尺寸自動調整卡片佈局（grid-cols-2 sm:grid-cols-4）
- **觸控友善** - 優化的觸控交互體驗

### 🗑️ 移除內容

- 移除所有 Python 腳本（generate_weekly_report.py 等）
- 移除週報告和月報告圖片檔案（reports/ 目錄）
- 移除 Excel 檔案和 CSV 資料檔（weekly/ 目錄）
- 移除 requirements.txt 和 Python 環境依賴
- 簡化專案結構為單一 HTML 應用

### 🔄 破壞性變更

- **從 Python 轉為 Web** - 不再需要 Python 環境
- **資料來源變更** - 直接從 Google Sheets 讀取而非本地檔案
- **報告格式變更** - 從靜態 Markdown 報告轉為互動式 Web 儀表板

### 📦 依賴套件

- Tailwind CSS 3.4.0 - 實用優先的 CSS 框架
- Chart.js 4.4.0 - 互動式圖表庫
- chartjs-adapter-date-fns 3.0.0 - 時間軸適配器
- chartjs-plugin-zoom 2.0.1 - 圖表縮放功能
- PapaParse 5.4.1 - CSV 解析
- jsPDF 2.5.1 - PDF 生成
- html2canvas 1.4.1 - HTML 截圖
- Font Awesome 6.4.0 - 圖示庫

---

## [1.x.x] - 2025-10-13

### Legacy Python 版本

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
