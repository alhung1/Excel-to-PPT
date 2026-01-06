# 📊 Excel to PowerPoint Generator

將 Excel 圖表和工作表自動轉換成 PowerPoint 簡報的網頁應用程式。

![Version](https://img.shields.io/badge/version-4.1.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.8+-green)

## ✨ 功能特色

- 🖱️ **拖放上傳** - 支援拖放 Excel 和 PPT 檔案
- 📊 **自動擷取** - 自動識別 Excel 中的圖表和工作表
- 🎯 **彈性對應** - 自由選擇要插入的頁面位置
- ⚙️ **可調整尺寸** - 自訂圖片在投影片中的位置和大小
- 📥 **即時下載** - 產生後立即下載 PowerPoint 檔案
- 🔄 **多檔案支援** - 可同時處理多個 Excel 檔案

## 🖥️ 系統需求

- **作業系統**: Windows 10/11 (需要 Windows COM 自動化)
- **Python**: 3.8 或更高版本
- **Microsoft Excel**: 必須安裝 (用於擷取圖表)

## 📦 安裝步驟

### 方法一：使用安裝腳本 (推薦)

```bash
# 執行安裝腳本
setup.bat
```

### 方法二：手動安裝

```bash
# 1. 建立虛擬環境
python -m venv venv

# 2. 啟動虛擬環境
venv\Scripts\activate

# 3. 安裝依賴套件
pip install -r requirements.txt
```

## 🚀 啟動服務

### 方法一：使用啟動腳本 (推薦)

```bash
# 雙擊執行
start_server.bat
```

### 方法二：手動啟動

```bash
# 啟動虛擬環境
venv\Scripts\activate

# 啟動服務
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

服務啟動後，開啟瀏覽器訪問：**http://localhost:8000**

## 📖 使用教學

### 步驟 1：上傳 Excel 檔案
- 將 Excel 檔案 (.xlsx, .xlsm, .xls) 拖放到左側上傳區
- 系統會自動讀取工作表和圖表資訊

### 步驟 2：上傳 PPT 模板
- 將 PowerPoint 檔案 (.pptx) 拖放到中間上傳區
- 可以看到模板的頁面列表

### 步驟 3：設定對應關係
1. 在左側列表中選擇要擷取的項目 (點擊選取)
2. 輸入目標頁碼
3. 點擊「➕ 加入」按鈕

### 步驟 4：調整圖片位置 (選用)
- **左**: 圖片左邊緣距離 (英吋)
- **上**: 圖片上邊緣距離 (英吋)
- **寬**: 圖片寬度 (英吋)
- **高**: 圖片高度 (英吋)

### 步驟 5：產生 PowerPoint
1. 輸入輸出檔名
2. 點擊「⚡ 產生 PowerPoint」
3. 等待處理完成後下載

## 🔧 API 文件

### 上傳 Excel 檔案
```
POST /api/upload-excel
Content-Type: multipart/form-data

Response:
{
  "status": "success",
  "file_id": "abc12345",
  "filename": "data.xlsx",
  "worksheets": [...],
  "chartsheets": [...]
}
```

### 上傳 PPT 模板
```
POST /api/upload-ppt
Content-Type: multipart/form-data

Response:
{
  "status": "success",
  "file_id": "xyz67890",
  "filename": "template.pptx",
  "total_slides": 12,
  "slides": [...]
}
```

### 產生 PowerPoint
```
POST /api/generate
Content-Type: application/json

Body:
{
  "template_id": "xyz67890",
  "output_name": "Report",
  "mappings": [
    {
      "excel_id": "abc12345",
      "name": "Sheet1",
      "page": 1,
      "type": "worksheet"
    }
  ],
  "img_left": 0.423,
  "img_top": 1.4,
  "img_width": 12.0,
  "img_height": 5.6
}

Response:
{
  "status": "success",
  "job_id": "job123",
  "download_url": "/api/download/job123/Report.pptx",
  "results": [...]
}
```

### 下載檔案
```
GET /api/download/{job_id}/{filename}
```

### 移除檔案
```
DELETE /api/remove-file/{file_id}
```

## 📁 專案結構

```
Excel-to-ppt/
├── app/
│   ├── main.py                    # FastAPI 主程式
│   ├── ppt_generator.py           # PPT 生成模組
│   ├── excel_chart_extractor.py   # Excel 圖表擷取模組
│   └── ...
├── static/
│   └── index.html                 # 前端頁面
├── uploads/                       # 上傳檔案暫存
├── outputs/                       # 輸出檔案
├── requirements.txt               # Python 依賴套件
├── start_server.bat              # 啟動腳本
├── setup.bat                     # 安裝腳本
└── README.md                     # 說明文件
```

## 🛠️ 技術細節

### 依賴套件

| 套件 | 用途 |
|------|------|
| FastAPI | Web 框架 |
| uvicorn | ASGI 伺服器 |
| python-pptx | PowerPoint 操作 |
| pywin32 | Windows COM 自動化 |
| openpyxl | Excel 讀取 |
| Pillow | 圖片處理 |

### 圖表擷取機制

本專案使用 Windows COM 自動化 (pywin32) 來擷取 Excel 圖表：

1. **Chart Sheet**: 直接使用 `Chart.Export()` 匯出
2. **Embedded Chart**: 透過 `ChartObjects().Chart.Export()` 匯出
3. **Worksheet (無圖表)**: 使用 `CopyPicture()` + 臨時圖表方式匯出

### 圖片驗證機制

為確保擷取品質，程式會：
- 驗證匯出檔案是否存在
- 檢查檔案大小是否合理 (> 500 bytes)
- **使用 PIL 檢查圖片內容** - 偵測空白/單色圖片
- 對失敗的擷取進行重試 (最多 3 次)
- 在剪貼簿操作間加入延遲避免競爭條件

### 多層備用擷取機制

當直接匯出失敗時，程式會自動嘗試備用方法：

```
1. Chart.Export() 直接匯出
   ↓ 失敗時
2. ChartObject.CopyPicture() + 貼到臨時 Chart Sheet
   ↓ 失敗時  
3. UsedRange.CopyPicture() 擷取整個使用範圍
```

每個方法都會驗證輸出圖片是否有效（非空白）。

## 🐛 已知問題與解決方案

### 問題：部分圖片無法成功轉換

**原因**：
- 剪貼簿被其他程式占用
- Excel 尚未完成匯出操作
- 工作表內容為空

**解決方案** (v4.0 已修復)：
- ✅ 加入圖片驗證機制
- ✅ 加入重試邏輯 (最多 3 次)
- ✅ 清除剪貼簿避免舊資料干擾
- ✅ 加入操作延遲確保完成

### 問題：圖表匯出後為空白圖片

**原因**：
- `Chart.Export()` 在某些情況下會產生有效大小但內容空白的圖片
- 嵌入圖表的直接匯出可能不穩定

**解決方案** (v4.1 已修復)：
- ✅ 使用 PIL 檢查圖片內容（獨特顏色數、標準差）
- ✅ 加入 `CopyPicture` 備用擷取方法
- ✅ 改用 Chart Sheet 作為臨時貼上目標
- ✅ 多層 fallback 機制確保成功率

### 問題：服務無法啟動

**解決方案**：
1. 確認已安裝 Python 3.8+
2. 確認已安裝 Microsoft Excel
3. 執行 `setup.bat` 重新安裝依賴

### 問題：Excel 檔案讀取失敗

**解決方案**：
1. 確認 Excel 檔案沒有被其他程式開啟
2. 確認檔案格式為 .xlsx, .xlsm 或 .xls
3. 嘗試用 Excel 開啟並重新儲存檔案

## 📝 更新日誌

### v4.1.0 (2026-01-06)
- 🐛 修復嵌入圖表匯出為空白圖片的問題
- ✨ 加入 PIL 圖片內容驗證（偵測空白/單色圖片）
- ✨ 加入多層備用擷取機制（CopyPicture fallback）
- ✨ 改用 Chart Sheet 作為臨時貼上目標（更穩定）
- 📊 加入詳細診斷日誌，方便追蹤問題

### v4.0.0 (2026-01-06)
- 🐛 修復部分圖片無法成功轉換的問題
- ✨ 加入圖片驗證機制
- ✨ 加入重試邏輯和延遲處理
- ✨ 改進錯誤訊息顯示
- 📊 加入擷取進度診斷資訊

### v3.0.0
- 🎨 全新 UI 設計
- ✨ 支援多 Excel 檔案
- ✨ 支援拖放上傳

## 📄 授權

MIT License

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request！

---

Made with ❤️ for productivity
