# 🏥 GLP-1 體重追蹤系統 v2.0

> 現代化醫療風格的體重管理儀表板，提供完整的身體組成分析與視覺化工具

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-2.0.0-brightgreen.svg)](https://github.com/RexhuangTW/GLP1-weight-tracking)
[![Status](https://img.shields.io/badge/status-active-success.svg)]()

---

## 📖 專案簡介

GLP-1 體重追蹤系統是一個完全基於 Web 的單頁應用程式，專為使用 GLP-1 藥物進行體重管理的使用者設計。本系統提供專業的醫療風格介面，整合 Google Sheets 資料來源，實現即時的體重、體脂、肌肉率和內臟脂肪追蹤與分析。

### ✨ 核心特色

- 🎨 **專業醫療風格** - 簡潔的藍白配色，符合醫療專業形象
- 📊 **互動式圖表** - 支援縮放、拖曳、即時篩選的視覺化分析
- 📱 **完全響應式** - 手機、平板、桌面全平台優化
- 💾 **離線支援** - Service Worker 實現離線瀏覽功能
- 🔔 **每日提醒** - 可設定每日推送通知提醒記錄體重
- 📥 **多格式匯出** - 支援 CSV 和 PDF 報告匯出
- 🌓 **深色模式** - 自動儲存的主題偏好設定

---

## 🚀 快速開始

### 線上使用

直接開啟 `index.html` 檔案即可使用，或部署到任何靜態網站託管服務：

```bash
# 使用 Python 本地測試
python3 -m http.server 8000

# 或使用 Node.js
npx serve

# 瀏覽器開啟
open http://localhost:8000
```

### 資料來源設定

1. 準備 Google Sheets 試算表，包含以下欄位：
   - `date` - 日期（格式：YYYY-MM-DD）
   - `w_am` - 早上體重
   - `w_pm` - 晚上體重
   - `f_am` - 早上體脂率
   - `m_am` - 早上肌肉率
   - `v` - 內臟脂肪

2. 將試算表發佈為 CSV 格式：
   - 檔案 → 共用 → 發佈到網路
   - 選擇「逗號分隔值 (.csv)」
   - 複製發佈的 URL

3. 在應用程式中設定：
   - 點擊左側選單的「設定」
   - 將 CSV URL 貼上到「資料來源 URL」欄位
   - 點擊「儲存設定」

---

## 📊 功能說明

### Phase 1 - 核心功能

#### 📥 CSV 資料匯出
一鍵匯出所有歷史數據為標準 CSV 格式，方便進行後續分析或備份。

#### 📋 表格筆數選擇器
彈性顯示表格資料：
- 15 筆 - 快速瀏覽最近紀錄
- 30 筆 - 一個月的數據
- 50 筆 - 詳細歷史
- 全部 - 完整資料集

#### 🌓 深色模式
支援深色與淺色主題切換，自動儲存使用者偏好，保護眼睛降低疲勞。

#### 📈 週比較分析
自動計算本週與上週的數據對比：
- 體重變化趨勢
- 體脂率增減
- 肌肉率變化
- 內臟脂肪對比
- 視覺化增減指標（🟢 減少 / 🔴 增加）

### Phase 2 - 進階功能

#### 🎯 進階目標設定
除了體重目標外，新增：
- **體脂率目標** - 設定理想體脂範圍
- **BMI 目標** - 設定目標 BMI 值
- 即時顯示與目標的差距

#### 🏆 最佳紀錄統計
自動追蹤並顯示：
- 💚 **最低體重** - 歷史最佳紀錄
- 🔴 **最高體重** - 起始點參考
- 💛 **最低體脂率** - 體脂管理成果
- 💪 **最高肌肉率** - 肌肉量巔峰

#### 📊 波動分析
深入分析數據穩定性：
- **體重波動度** - 標準差計算
- **體脂波動度** - 變化趨勢分析
- **最大單日變化** - 異常值偵測
- **最長連降紀錄** - 持續進步天數

#### 📱 響應式優化
大幅改善手機版使用體驗：
- 2 欄網格佈局（手機）
- 4 欄網格佈局（桌面）
- 自適應字體大小
- 觸控優化的互動元素

### Phase 3 - 高級功能

#### 📄 PDF 報告匯出
使用 jsPDF 和 html2canvas 生成專業 PDF 報告：
- 完整的視覺化圖表
- 統計數據摘要
- 自動分頁處理
- 高解析度輸出

#### ✏️ 數據輸入/編輯
本地數據管理功能：
- 新增體重紀錄
- 編輯現有數據
- LocalStorage 儲存
- 與 Google Sheets 資料合併顯示

#### 🔔 每日提醒
設定每日推送通知：
- 自訂提醒時間
- 瀏覽器通知 API
- 持久化設定
- 一鍵開關

#### 📴 離線支援
Service Worker 實現離線功能：
- 離線瀏覽歷史數據
- 快取關鍵資源
- 自動更新機制
- 漸進式 Web App

#### 🔍 圖表互動增強
Chart.js 進階功能：
- **滾輪縮放** - 精確查看數據細節
- **拖曳平移** - 瀏覽不同時間區段
- **雙擊重置** - 快速回到預設視圖
- **一鍵重置** - 重置所有圖表縮放

---

## 🛠️ 技術架構

### 前端技術棧

| 技術 | 版本 | 用途 |
|------|------|------|
| **Tailwind CSS** | 3.4.0 | 實用優先的 CSS 框架 |
| **Chart.js** | 4.4.0 | 互動式圖表繪製 |
| **chartjs-adapter-date-fns** | 3.0.0 | 時間軸格式化 |
| **chartjs-plugin-zoom** | 2.0.1 | 圖表縮放功能 |
| **PapaParse** | 5.4.1 | CSV 解析 |
| **jsPDF** | 2.5.1 | PDF 文件生成 |
| **html2canvas** | 1.4.1 | HTML 轉圖片 |
| **Font Awesome** | 6.4.0 | 圖示庫 |

### 核心特性

- ✅ **單一檔案架構** - 2855 行完整功能
- ✅ **無需建置工具** - 直接開啟即可使用
- ✅ **CDN 載入** - 所有依賴透過 CDN
- ✅ **LocalStorage** - 30 分鐘快取機制
- ✅ **Service Worker** - 離線支援
- ✅ **Notification API** - 推送通知
- ✅ **Google Sheets API** - 即時資料同步

### 資料流程

```
Google Sheets (CSV 發佈)
    ↓
PapaParse 解析
    ↓
LocalStorage 快取 (30 分鐘)
    ↓
資料處理 & 計算
    ↓
Chart.js 視覺化
    ↓
使用者互動 (縮放/篩選/匯出)
```

---

## 📁 專案結構

```
GLP1-weight-tracking/
│
├── index.html          # 主應用程式（2855 行）
├── CHANGELOG.md        # 更新日誌
├── README.md           # 專案說明（本檔案）
└── LICENSE             # MIT 授權
```

---

## 🎨 設計系統

### 顏色配置

| 用途 | 顏色代碼 | Tailwind 類別 |
|------|----------|--------------|
| **主色（藍色）** | `#2563EB` | `blue-600` |
| **次色（青色）** | `#0891B2` | `cyan-600` |
| **強調色（綠色）** | `#059669` | `emerald-600` |
| **警示色（紅色）** | `#DC2626` | `red-600` |
| **背景色** | `#F8FAFC` | `slate-50` |
| **卡片背景** | `#FFFFFF` | `white` |
| **文字色** | `#1E293B` | `slate-800` |

### 醫療風格元素

- 📊 **圓角卡片** - `rounded-lg` 柔和邊角
- 🎯 **線條圖示** - Font Awesome Regular (far)
- 💫 **淡入動畫** - `fade-in` 簡約過場
- 🏥 **InBody 視覺化** - 專業人體成分圖
- 📐 **網格佈局** - 整齊的資料排列

---

## 📱 螢幕截圖

### 桌面版
![Desktop View](https://via.placeholder.com/800x400?text=Desktop+View)

### 手機版
![Mobile View](https://via.placeholder.com/400x800?text=Mobile+View)

### 深色模式
![Dark Mode](https://via.placeholder.com/800x400?text=Dark+Mode)

---

## 🔐 隱私與安全

- ✅ **本地儲存** - 所有設定儲存在瀏覽器 LocalStorage
- ✅ **唯讀資料** - 僅讀取 Google Sheets，不寫入
- ✅ **無伺服器** - 完全前端運作，無後端追蹤
- ✅ **離線優先** - 資料快取在本地，減少網路請求
- ✅ **開源透明** - 完整程式碼可檢視與審核

---

## 📈 未來計畫

### v2.1 計畫功能

- [ ] 多語言支援（英文、簡中）
- [ ] 匯入/匯出完整設定
- [ ] 自訂圖表顏色
- [ ] 更多統計指標（BMR、TDEE）
- [ ] 資料比較模式（月對月、年對年）

### v3.0 願景

- [ ] 雲端資料同步（Firebase / Supabase）
- [ ] 多使用者支援
- [ ] 社群分享功能
- [ ] AI 趨勢預測
- [ ] 營養建議整合
- [ ] 運動記錄追蹤

---

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request！

### 開發流程

1. Fork 本專案
2. 建立功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

### 程式碼風格

- 使用 Tailwind CSS 類別
- 遵循 ES6+ JavaScript 標準
- 保持程式碼簡潔易讀
- 添加必要的註解

---

## 📄 授權

本專案採用 MIT 授權 - 詳見 [LICENSE](LICENSE) 檔案

---

## 👨‍💻 作者

**Rex Huang**

- GitHub: [@RexhuangTW](https://github.com/RexhuangTW)
- Email: your.email@example.com

---

## 🙏 致謝

- [Chart.js](https://www.chartjs.org/) - 強大的圖表庫
- [Tailwind CSS](https://tailwindcss.com/) - 實用的 CSS 框架
- [Font Awesome](https://fontawesome.com/) - 豐富的圖示集
- [PapaParse](https://www.papaparse.com/) - 快速的 CSV 解析器
- [jsPDF](https://github.com/parallax/jsPDF) - 客戶端 PDF 生成

---

## 📞 聯絡資訊

如有任何問題或建議，歡迎透過以下方式聯絡：

- 📧 Email: your.email@example.com
- 💬 GitHub Issues: [提交 Issue](https://github.com/RexhuangTW/GLP1-weight-tracking/issues)
- 🐛 Bug Report: [回報錯誤](https://github.com/RexhuangTW/GLP1-weight-tracking/issues/new?template=bug_report.md)
- 💡 Feature Request: [建議功能](https://github.com/RexhuangTW/GLP1-weight-tracking/issues/new?template=feature_request.md)

---

<div align="center">

**⭐️ 如果這個專案對你有幫助，請給個星星支持！ ⭐️**

Made with ❤️ by Rex Huang

</div>
