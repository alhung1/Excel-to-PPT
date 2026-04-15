# 📊 Excel to PowerPoint Generator

將 Excel 圖表和工作表自動轉換成 PowerPoint 簡報的網頁應用程式。

![Version](https://img.shields.io/badge/version-6.0.0-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Python](https://img.shields.io/badge/python-3.8+-green)

## ✨ 功能特色

- 🖱️ **拖放上傳** - 支援拖放 Excel 和 PPT 檔案
- 📊 **自動擷取** - 自動識別 Excel 中的圖表和工作表
- 🎯 **彈性對應** - 自由選擇要插入的頁面位置
- ⚙️ **可調整尺寸** - 自訂圖片在投影片中的位置和大小
- 📥 **即時下載** - 產生後立即下載 PowerPoint 檔案
- 🔄 **多檔案支援** - 可同時處理多個 Excel 檔案
- ✏️ **可編輯圖表** - 支援插入可編輯的圖表，每個項目可獨立選擇模式
- 🩺 **健康檢查 API** - `/api/health` 端點監控服務狀態 (v6.0 新增)
- 🧹 **自動清理** - 定時清理過期的上傳與輸出檔案 (v6.0 新增)
- 📝 **結構化日誌** - Python logging 模組取代 print (v6.0 新增)

## 🖥️ 系統需求

- **作業系統**: Windows 10/11 (需要 Windows COM 自動化)
- **Python**: 3.8 或更高版本
- **Microsoft Excel**: 必須安裝 (用於擷取圖表)
- **Microsoft PowerPoint**: 必須安裝 (可編輯圖表模式需要)

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
2. **選擇圖表模式** - 每個項目可獨立選擇：
   - **圖片**: 以靜態圖片插入，無法在 PowerPoint 中編輯
   - **可編輯**: 以可編輯物件插入，可在 PowerPoint 中修改資料、格式
3. 輸入目標頁碼
4. 點擊「➕ 加入」按鈕

> 💡 **提示**: 可以在同一份 PPT 中混合使用「圖片」和「可編輯」兩種模式。可編輯圖表模式需要 Excel 和 PowerPoint 在處理時短暫顯示視窗。

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
      "type": "worksheet",
      "chart_mode": "image"
    },
    {
      "excel_id": "abc12345",
      "name": "Chart1",
      "page": 2,
      "type": "chartsheet",
      "chart_mode": "embedded"
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
  "results": [...],
  "mode": "mixed"
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

### 健康檢查 (v6.0 新增)
```
GET /api/health

Response:
{
  "status": "ok",
  "version": "6.0.0",
  "uploads_count": 3,
  "outputs_dir_size_mb": 45.2
}
```

## 📁 專案結構

```
Excel-to-ppt/
├── app/
│   ├── main.py                       # FastAPI 應用初始化 + middleware
│   ├── config.py                     # 配置常數、logging、路徑
│   ├── models/
│   │   └── schemas.py                # Pydantic 資料模型
│   ├── routers/
│   │   └── api.py                    # API 端點定義
│   ├── services/
│   │   ├── excel_service.py          # Excel COM 操作 (含 context manager)
│   │   ├── ppt_service.py            # PPT 生成邏輯
│   │   └── file_manager.py           # 檔案管理與定時清理
│   └── utils/
│       ├── image_validator.py        # 圖片驗證 (PIL 檢查)
│       └── clipboard.py             # 剪貼簿操作工具
├── cli/
│   └── report_cli.py                # 命令列報告產生器 (統一版)
├── tools/
│   ├── find_charts.py               # Excel 圖表探索工具
│   └── analyze_ppt.py               # PPT 版面分析工具
├── static/
│   └── index.html                    # 前端頁面
├── templates/                        # PPT 模板
├── uploads/                          # 上傳檔案暫存
├── outputs/                          # 輸出檔案
├── tests/
│   └── test_refactored.py           # 測試套件 (40 項測試)
├── requirements.txt                  # Python 依賴套件
├── start_server.bat                 # Windows 啟動腳本
├── start_server.sh                  # Linux/Mac 啟動腳本
├── setup.bat                        # 安裝腳本
├── stop_server.bat                  # 停止腳本
└── README.md                        # 說明文件
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
| Pillow | 圖片處理與驗證 |
| pydantic | 資料驗證 |
| httpx | HTTP 測試客戶端 |

### 圖表模式說明

| 模式 | 說明 | 優點 | 缺點 |
|------|------|------|------|
| **圖片模式** | 圖表匯出為 PNG 圖片後插入 | 處理快速、相容性高 | 無法在 PPT 中編輯 |
| **可編輯模式** | 圖表以 Office 物件複製貼上 | 可在 PPT 中編輯資料和格式 | 需要 PPT 安裝、處理稍慢 |

### v6.0 架構改善

- **模組化設計**: main.py 從 1020+ 行精簡為 ~70 行，邏輯拆分為 8 個獨立模組
- **COM Context Manager**: `ExcelCOM` / `PowerPointCOM` 確保 COM 物件必定釋放，防止進程洩漏
- **非阻塞 API**: `generate` 端點改為 sync `def`，FastAPI 自動用執行緒池處理，不再阻塞事件迴圈
- **結構化日誌**: Python `logging` 模組取代 `print()`，支援分級和時間戳
- **自動檔案清理**: 透過 FastAPI lifespan 定時清理超過 24 小時的暫存檔案
- **統一 CLI 工具**: 三個 netgear_report 腳本合併為 `cli/report_cli.py`，支援 config/interactive/direct 三種模式

### CLI 命令列工具

```bash
# 使用 JSON 設定檔
python -m cli.report_cli --config report_config.json

# 互動模式
python -m cli.report_cli --interactive --excel data.xlsm --template report.pptx

# 直接指定 mapping
python -m cli.report_cli --excel data.xlsm --template report.pptx \
    --output result.pptx \
    --map "Metric DUT vs REF#1:8:worksheet" \
    --map "BI:9:chartsheet"
```

### 圖表擷取機制

使用 Windows COM 自動化 (pywin32) 擷取 Excel 圖表：

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

## 🧪 測試

```bash
# 執行測試套件 (40 項測試)
python tests/test_refactored.py
```

測試涵蓋：模組匯入、配置值、Pydantic 模型驗證、圖片驗證器、檔案管理器、PPT 服務函式、FastAPI 端點。

## 🐛 已知問題與解決方案

### 問題：部分圖片無法成功轉換

**原因**: 剪貼簿被其他程式占用、Excel 尚未完成匯出、工作表內容為空

**解決方案**: 加入圖片驗證 + 重試邏輯 + 多層 fallback 機制

### 問題：圖表匯出後為空白圖片

**原因**: `Chart.Export()` 在某些情況下產生有效大小但內容空白的圖片

**解決方案**: 使用 PIL 檢查圖片內容 + CopyPicture 備用擷取 + 臨時 Chart Sheet

### 問題：服務無法啟動

**解決方案**:
1. 確認已安裝 Python 3.8+
2. 確認已安裝 Microsoft Excel
3. 執行 `setup.bat` 重新安裝依賴

## 📝 更新日誌

### v6.0.0 (2026-04-05)
- ♻️ **全面架構重構** - main.py 從 1020+ 行拆分為 8 個獨立模組
- 🛡️ **COM Context Manager** - ExcelCOM/PowerPointCOM 防止進程洩漏
- ⚡ **修復 async 阻塞** - generate 端點不再阻塞事件迴圈
- 📝 **結構化日誌** - Python logging 取代 print()
- 🧹 **自動檔案清理** - 定時清理過期暫存檔案
- 🩺 **健康檢查 API** - 新增 /api/health 端點
- 🔧 **統一 CLI 工具** - 三個 netgear_report 合併為 cli/report_cli.py
- ✅ **測試套件** - 40 項自動化測試

### v5.1.0 (2026-01-29)
- ✨ **獨立圖表模式選擇** - 每個對應項目可各自選擇圖片或可編輯模式
- ✨ 支援混合模式 - 同一份 PPT 可同時包含圖片和可編輯圖表

### v5.0.0 (2026-01-23)
- ✨ **新增可編輯圖表模式** - 圖表可在 PowerPoint 中編輯資料和格式
- ✨ 使用 PowerPoint COM 自動化實現可編輯圖表貼上

### v4.1.0 (2026-01-06)
- 🐛 修復嵌入圖表匯出為空白圖片的問題
- ✨ 加入 PIL 圖片內容驗證和多層備用擷取機制

### v4.0.0 (2026-01-06)
- 🐛 修復部分圖片無法成功轉換的問題
- ✨ 加入圖片驗證機制和重試邏輯

### v3.0.0
- 🎨 全新 UI 設計
- ✨ 支援多 Excel 檔案和拖放上傳

## 📄 授權

MIT License

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request！

---

Made with ❤️ for productivity
