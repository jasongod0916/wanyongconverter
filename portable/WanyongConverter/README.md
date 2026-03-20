# 萬用轉檔王 Portable

這是萬用轉檔王的 portable 版本。

你不需要另外安裝這套軟體，只要保留整個資料夾，就可以直接在 Windows 電腦上使用。

## 如何使用

1. 開啟 `WanyongConverter.exe`
2. 選擇要轉換的檔案
3. 選擇輸出格式
4. 開始轉檔

## 重要提醒

- 請不要只單獨搬走 `WanyongConverter.exe`
- 請保留整個 `WanyongConverter` 資料夾
- `tools` 資料夾裡的工具是這個 portable 版本的一部分

## 目前支援

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
  - 有 Microsoft Word 時：優先使用 Word 進行高還原輸出
  - 沒有 Microsoft Word 時：自動改走 portable fallback 路線
- `doc -> pdf`
  - 需要 Microsoft Word

### 影音

- 透過 ffmpeg 支援常見影片與音訊格式互轉

## 內含工具

- `tools\ffmpeg\ffmpeg.exe`
- `tools\pandoc\pandoc.exe`

## 移除方式

這個版本不需要安裝。

如果你不想用了，直接刪掉整個 `WanyongConverter` 資料夾即可。
