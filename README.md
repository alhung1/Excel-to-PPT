# Excel to PowerPoint Generator

從 Excel 擷取圖表，自動貼到 PowerPoint 模板中。

---

## 🚀 快速開始

### 方法一：使用啟動腳本（推薦）

| 步驟 | 動作 |
|------|------|
| **首次使用** | 雙擊 `setup.bat` → 安裝依賴套件（只需執行一次） |
| **啟動服務** | 雙擊 `start_server.bat` → 自動開啟瀏覽器 |
| **停止服務** | 按 `Ctrl+C` 或關閉視窗，或雙擊 `stop_server.bat` |
| **建立捷徑** | 雙擊 `create_shortcut.bat` → 桌面出現捷徑圖示 |

### 方法二：手動啟動

```powershell
cd C:\Users\alhung\excel-to-ppt
.\venv\Scripts\Activate.ps1
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

### 開啟網頁

瀏覽器打開：**http://localhost:8000**

---

## 📁 檔案說明

| 檔案 | 用途 |
|------|------|
| `start_server.bat` | 🚀 **主要啟動腳本** - 雙擊啟動服務 + 自動開瀏覽器 |
| `stop_server.bat` | 🛑 停止服務 |
| `setup.bat` | 🔧 首次安裝/重新安裝依賴套件 |
| `create_shortcut.bat` | 📌 在桌面建立捷徑 |

---

## 📖 操作流程

### Step 1️⃣ 上傳 Excel 檔案

- **拖曳** Excel 檔案到上傳區域
- 或 **點擊** 上傳區域選擇檔案
- 可上傳 **多個** Excel 檔案
- 支援 `.xlsx`, `.xlsm`, `.xls`

### Step 2️⃣ 上傳 PPT 模板

- **拖曳** PPT 檔案到上傳區域
- 或 **點擊** 上傳區域選擇檔案
- 支援 `.pptx`, `.ppt`

### Step 3️⃣ 設定圖表對應

1. **點擊** 左側列表中的項目（會變綠色表示選取）
2. **輸入** 要貼到的頁碼
3. **點擊**「➕ 加入」按鈕
4. 重複以上步驟加入更多對應

### Step 4️⃣ 調整設定（選填）

可調整圖片在 PPT 中的位置和大小：
- **左**: 左邊距 (預設 0.423 英寸)
- **上**: 上邊距 (預設 1.1 英寸)
- **寬**: 圖片寬度 (預設 12.0 英寸)
- **高**: 圖片高度 (預設 5.6 英寸)

### Step 5️⃣ 產生並下載

1. 輸入輸出檔案名稱
2. 點擊「⚡ 產生 PowerPoint」
3. 點擊「📥 下載 PowerPoint」

---

## 🎯 功能特色

| 功能 | 說明 |
|------|------|
| 📤 拖曳上傳 | 直接拖曳檔案到網頁 |
| 📁 多 Excel 支援 | 可同時載入多個 Excel 檔案 |
| 📊 自動擷取圖表 | 支援 Worksheet 和 Chart Sheet |
| 🎨 自訂位置大小 | 調整圖片在 PPT 中的呈現 |
| 🔄 即時預覽 | 查看 PPT 頁面標題 |

---

## 📋 圖表類型說明

| 類型 | 說明 | 標示 |
|------|------|------|
| Worksheet | 一般工作表（會擷取表格或內嵌圖表） | `sheet` |
| Chart Sheet | 獨立圖表頁（整頁都是圖表） | `chart` |

---

## ❓ 常見問題

### Q: 伺服器無法啟動 (Port 8000 被佔用)

```powershell
# 方法一：使用停止腳本
雙擊 stop_server.bat

# 方法二：手動關閉
Get-NetTCPConnection -LocalPort 8000 | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

### Q: 上傳後看不到圖表

- 確認 Excel 檔案中有圖表（不是圖片）
- Chart Sheet 會顯示為 `chart` 標籤
- Worksheet 中的內嵌圖表也會被擷取

### Q: 圖片大小不對

調整右側的位置設定：
- 左/上：調整圖片位置
- 寬/高：調整圖片大小

### Q: 套件安裝失敗

```powershell
# 重新執行安裝
雙擊 setup.bat
```

---

## 📁 專案結構

```
excel-to-ppt/
├── app/
│   ├── main.py                  # FastAPI 主程式
│   ├── ppt_generator.py         # PPT 生成器
│   └── excel_chart_extractor.py # Excel 圖表擷取
├── static/
│   └── index.html               # 網頁介面
├── uploads/                     # 上傳的檔案
├── outputs/                     # 產生的 PPT
├── venv/                        # Python 虛擬環境
├── start_server.bat             # 啟動腳本
├── stop_server.bat              # 停止腳本
├── setup.bat                    # 安裝腳本
├── create_shortcut.bat          # 建立桌面捷徑
├── requirements.txt
└── README.md
```

---

## 🔧 環境需求

- Windows 作業系統
- Microsoft Excel（已安裝）
- Python 3.10+

---

## 📝 手動安裝（進階）

如果 `setup.bat` 無法使用，可手動安裝：

```powershell
cd C:\Users\alhung\excel-to-ppt
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## 📄 License

MIT License
