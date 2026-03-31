# 萬用轉檔王

萬用轉檔王是一套 Windows 攜帶式轉檔工具，目標是讓圖片、文件、影音轉檔都能在同一個介面完成，而且整包可以直接帶走，不必額外安裝成正式系統軟體。

## 特色

- 單一介面處理圖片、文件、影音轉檔
- 提供 portable 版本，可直接整包複製到其他 Windows 電腦
- 支援中文檔名與中文路徑
- `docx -> pdf` 採雙模式
  - 有 Microsoft Word：優先使用 Word，高還原輸出
  - 沒有 Microsoft Word：自動改走 portable fallback 路線
- 內含 `ffmpeg`、`pandoc`，不需要另外安裝這兩套工具才能使用 portable 版

## 快速開始

### 一般使用者

如果你只是想直接使用軟體，建議從 GitHub `Releases` 下載 portable 壓縮包，不需要先安裝 Python，也不需要用 git clone 原始碼。

直接開啟：

```text
portable\WanyongConverter\WanyongConverter.exe
```

使用方式：

1. 選擇要轉換的檔案
2. 選擇輸出格式
3. 開始批次轉檔
4. 到輸出資料夾查看結果

如果要分享給別人，請整個 `portable\WanyongConverter` 資料夾一起提供，不要只單獨拿 `exe`。

### 開發者

直接執行原始碼：

```powershell
python .\super_converter.py
```

或使用 GUI 啟動入口：

```powershell
python .\launch_super_converter.pyw
```

## 支援格式

### 圖片

- 目前實測來源格式：`png`、`jpg`、`bmp`、`gif`、`tiff`、`webp`、`ico`、`svg`
- 目前實測成功的圖片輸出格式：
  `apng`, `avif`, `avifs`, `bmp`, `bw`, `dds`, `dib`, `gif`, `icb`, `icns`, `ico`, `im`
  `j2c`, `j2k`, `jfif`, `jp2`, `jpc`, `jpe`, `jpeg`, `jpf`, `jpg`, `jpx`, `mpo`, `msp`
  `pbm`, `pcx`, `pdf`, `pfm`, `pgm`, `png`, `pnm`, `ppm`, `qoi`, `rgb`, `rgba`, `sgi`
  `tga`, `tif`, `tiff`, `vda`, `vst`, `webp`, `xbm`
- 以上圖片格式組合已完成本機矩陣測試，最近一次結果為 `331 / 331` 成功

### 文件

- 目前實測來源格式：`doc`、`docbook`、`docx`、`epub`、`htm`、`html`、`latex`、`markdown`、`md`、`odt`、`org`、`pdf`、`rst`、`rtf`、`tex`、`txt`
- `docx -> md`
- `md -> docx`
- `md -> pdf`
- `pdf -> md`
- `pdf -> docx`
- `pdf -> txt`
- `txt / md / html` 互轉
- `docx / odt / rtf / epub / md / html / txt` 可透過 pandoc 互轉
- `docx -> pdf`
  - 有 Word 時使用高還原模式
  - 沒有 Word 時使用 portable fallback 模式
- `doc -> pdf`
  - 需要 Microsoft Word
- 目前實測成功的文件輸出格式：
  - `docx`
  - `epub`
  - `html`
  - `md`
  - `odt`
  - `pdf`
  - `rtf`
  - `txt`
- 以上文件格式組合已完成本機矩陣測試，最近一次結果為 `109 / 109` 成功

### 影音

- 目前實測影片來源格式：`avi`、`flv`、`mkv`、`mov`、`mp4`、`webm`、`wmv`
- 目前實測音訊來源格式：`aac`、`flac`、`m4a`、`mp3`、`ogg`、`wav`
- 目前實測成功的影音輸出格式：
  - `mp4`
  - `mov`
  - `avi`
  - `mkv`
  - `webm`
  - `flv`
  - `wmv`
  - `mp3`
  - `wav`
  - `aac`
  - `flac`
  - `ogg`
  - `m4a`
  - `gif`
- 補充：`gif` 只會出現在有畫面的影片來源，不會出現在純音訊來源
- 以上影音格式組合已完成本機矩陣測試，最近一次結果為 `163 / 163` 成功

## 限制與說明

- `doc -> pdf` 依賴 Microsoft Word，沒有 Word 的電腦無法走這條高還原路線
- `docx -> pdf` 在沒有 Word 時仍可輸出 PDF，但版面不保證與原始 Word 完全一致
- 這個 repo 內含 portable 成品與大型執行檔，clone 與下載體積會比較大

## 核心工具

- `ffmpeg`
- `pandoc`
- `pdf2docx`
- `Microsoft Word`
  - 只在 `doc/docx -> pdf` 的高還原模式下使用

## 專案結構

```text
super_converter.py                主程式與轉檔邏輯
launch_super_converter.pyw        Windows GUI 啟動入口
build_portable.ps1                打包 portable 版本
build_exe.ps1                     其他打包腳本
tools\pandoc\pandoc.exe           打包使用的 pandoc
portable\WanyongConverter\        已打包完成的 portable 成品
```

## 打包 portable

```powershell
powershell -ExecutionPolicy Bypass -File .\build_portable.ps1
```

輸出位置：

```text
portable\WanyongConverter\
```

## portable 內容

```text
portable\WanyongConverter\WanyongConverter.exe
portable\WanyongConverter\tools\ffmpeg\ffmpeg.exe
portable\WanyongConverter\tools\pandoc\pandoc.exe
```

## 移除方式

portable 版本不需要安裝。刪掉整個 `portable\WanyongConverter` 資料夾即可移除。

## GitHub 備註

這個 repo 為了保留完整 portable 成品，使用 Git LFS 追蹤大型 `pandoc.exe` 檔案。
