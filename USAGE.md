# 快速使用指南

## 基本使用

### 1. 產生最新一週報告
```bash
python3 generate_weekly_report.py BodyComposition_202507-202510.csv
```

### 2. 產生總結報告
```bash
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary
```

### 3. 產生特定週次報告
```bash
# 產生第1週報告（2025-08-15 ~ 2025-08-21）
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index 1

# 產生第5週報告
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index 5
```

### 4. 批量產生所有週次報告
```bash
# 產生第1-9週的所有報告
for i in {1..9}; do
  python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index $i
done
```

## CSV 數據格式說明

從體脂計匯出的 CSV 檔案應包含：
- **測量日期**：格式為 `YYYY/MM/DD HH:MM`
- **體重(kg)**：體重數值
- **體脂肪(%)**：體脂率數值

腳本會自動：
1. 將測量時間分類為「早上」（5:00-12:00）和「晚上」（其他時間）
2. 若同一天同一時段有多次測量，會計算平均值
3. 按日期彙整數據並產生報告

## 週次計算方式

- **Anchor Date（基準日）**：2025-08-15（週四）
- **週次起算**：從 Anchor Date 開始，每7天為一週
  - 第1週：2025-08-15 ~ 2025-08-21
  - 第2週：2025-08-22 ~ 2025-08-28
  - 第3週：2025-08-29 ~ 2025-09-04
  - ...以此類推

## 輸出檔案說明

### 週報輸出
- `weekly/2025-CW01_weight_tracking.xlsx` - 該週的 Excel 數據
- `reports/2025-CW01/2025-CW01_weekly_report.md` - 該週的 Markdown 報告
- `reports/2025-CW01/2025-CW01_weight_trend.png` - 體重趨勢圖
- `reports/2025-CW01/2025-CW01_bodyfat_trend.png` - 體脂趨勢圖

### 總結報告輸出
- `reports/summary/overall_summary_report.md` - 總結報告
- `reports/summary/summary_weight_trend.png` - 總體體重趨勢圖
- `reports/summary/summary_bodyfat_trend.png` - 總體體脂趨勢圖

## 常用命令組合

### 更新 CSV 後的完整工作流程
```bash
# 1. 產生所有週次報告
for i in {1..9}; do
  python3 generate_weekly_report.py BodyComposition_202507-202510.csv --week-index $i
done

# 2. 產生總結報告
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary

# 3. 查看最新週報
cat reports/2025-CW09/2025-CW09_weekly_report.md
```

### 只更新最新數據
```bash
# 產生最新一週和總結報告
python3 generate_weekly_report.py BodyComposition_202507-202510.csv
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --summary
```

## 進階選項

### 自訂 Anchor Date
```bash
# 使用不同的起始日期
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --anchor-date 2025-09-01
```

### 指定輸出目錄
```bash
# 將報告輸出到其他目錄
python3 generate_weekly_report.py BodyComposition_202507-202510.csv --out-root /path/to/output
```

## 數據品質檢查

### 檢查缺失數據
週報中顯示 `nan` 表示該時段沒有測量數據，這是正常的。腳本會：
- 自動跳過缺失的數據點
- 在計算統計時只使用有效數據
- 在趨勢圖中顯示所有可用數據點

### 建議測量時間
- **早上**：起床後、上廁所後、早餐前（5:00-12:00）
- **晚上**：睡前或晚餐後2小時（12:00後到隔天5:00前）
- 保持每天相同時間測量以確保數據一致性
