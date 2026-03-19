# 萬用轉檔王

萬用轉檔王是 Windows 攜帶版轉檔工具，主打不用安裝、整包可帶走。

## 目前穩定支援

### 圖片
- 常見圖片格式互轉
- 圖片轉 PDF

### 文件
- `docx -> md`
- `md -> docx`
- `md -> pdf`
- `pdf -> md`
- `pdf -> docx`
- `pdf -> txt`
- `txt/md/html` 互轉
- `docx/odt/rtf/epub/md/html/txt` 之間可透過 pandoc 轉換

### 影音
- 透過 ffmpeg 支援常見影片與音訊格式互轉

## 目前不支援
- `.doc`
- 舊版 Office 大範圍互轉
- 高還原排版的 `docx -> pdf`

說明：現在的 `docx -> pdf` 是穩定 fallback 路線，重點是能轉成功，不保證和 Word 版面完全一致。

## 核心工具
- `ffmpeg`
- `pandoc`
- `pdf2docx`

## 執行
```powershell
python super_converter.py
```

## 打包攜帶版
```powershell
powershell -ExecutionPolicy Bypass -File .\build_portable.ps1
```

輸出會在：`portable\萬用轉檔王\`

## 攜帶版內容
- `portable\萬用轉檔王\萬用轉檔王.exe`
- `portable\萬用轉檔王\tools\ffmpeg\ffmpeg.exe`
- `portable\萬用轉檔王\tools\pandoc\pandoc.exe`

## 備註
- 這一版不再依賴其他 Office 後端
- 轉檔時會自動處理中文檔名與中文路徑
- 刪掉整個 `portable\萬用轉檔王` 資料夾就等於移除
