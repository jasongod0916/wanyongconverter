# 萬用轉檔王

萬用轉檔王是 Windows 用的攜帶式轉檔工具，支援圖片、文件、影音格式互轉，並提供可直接分享的 portable 版本。

## 功能

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
- `txt / md / html` 互轉
- `docx / odt / rtf / epub / md / html / txt` 可透過 pandoc 互轉
- `docx -> pdf`
  - 有 Microsoft Word 時：優先使用 Word 進行高還原匯出
  - 沒有 Microsoft Word 時：使用 portable fallback 路線
- `doc -> pdf`
  - 需要 Microsoft Word

### 影音
- 透過 ffmpeg 支援常見影片與音訊格式互轉

## 核心工具
- `ffmpeg`
- `pandoc`
- `pdf2docx`
- `Microsoft Word`（僅在 `doc/docx -> pdf` 高還原模式會使用）

## 目錄結構

- `super_converter.py`
  - 主要轉檔邏輯與 GUI
- `launch_super_converter.pyw`
  - Windows GUI 啟動入口
- `build_portable.ps1`
  - 產生 portable 版本
- `tools\pandoc\pandoc.exe`
  - portable 打包用 pandoc
- `portable\WanyongConverter`
  - 已打包完成的 portable 成品

## 執行

```powershell
python .\super_converter.py
```

或者：

```powershell
python .\launch_super_converter.pyw
```

## 打包 portable

```powershell
powershell -ExecutionPolicy Bypass -File .\build_portable.ps1
```

輸出會在：`portable\WanyongConverter\`

## portable 使用方式

1. 開啟 `portable\WanyongConverter\WanyongConverter.exe`
2. 要分享給別人時，整個 `WanyongConverter` 資料夾一起帶走
3. 刪掉整個 `portable\WanyongConverter` 資料夾即可移除

## GitHub 備註

這個 repo 包含 portable 成品與大型執行檔，因此 `pandoc.exe` 使用 Git LFS 追蹤。
