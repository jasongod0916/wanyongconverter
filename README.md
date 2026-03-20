# \u842c\u7528\u8f49\u6a94\u738b

\u842c\u7528\u8f49\u6a94\u738b\u662f Windows \u7528\u7684\u6536\u7d0d\u5f0f\u8f49\u6a94\u5de5\u5177\uff0c\u652f\u63f4\u5716\u7247\u3001\u6587\u4ef6\u3001\u5f71\u97f3\u683c\u5f0f\u4e92\u8f49\uff0c\u4e26\u63d0\u4f9b\u53ef\u76f4\u63a5\u5206\u4eab\u7684 portable \u7248\u672c\u3002

## \u529f\u80fd

### \u5716\u7247
- \u5e38\u898b\u5716\u7247\u683c\u5f0f\u4e92\u8f49
- \u5716\u7247\u8f49 PDF

### \u6587\u4ef6
- `docx -> md`
- `md -> docx`
- `md -> pdf`
- `pdf -> md`
- `pdf -> docx`
- `pdf -> txt`
- `txt / md / html` \u4e92\u8f49
- `docx / odt / rtf / epub / md / html / txt` \u53ef\u900f\u904e pandoc \u4e92\u8f49
- `docx -> pdf`
  - \u6709 Microsoft Word \u6642\uff1a\u512a\u5148\u4f7f\u7528 Word \u9032\u884c\u9ad8\u9084\u539f\u532f\u51fa
  - \u6c92\u6709 Microsoft Word \u6642\uff1a\u4f7f\u7528 portable fallback \u8def\u7dda
- `doc -> pdf`
  - \u9700\u8981 Microsoft Word

### \u5f71\u97f3
- \u900f\u904e ffmpeg \u652f\u63f4\u5e38\u898b\u5f71\u7247\u8207\u97f3\u8a0a\u683c\u5f0f\u4e92\u8f49

## \u6838\u5fc3\u5de5\u5177
- `ffmpeg`
- `pandoc`
- `pdf2docx`
- `Microsoft Word`\uff08\u50c5\u5728 `doc/docx -> pdf` \u9ad8\u9084\u539f\u6a21\u5f0f\u6703\u4f7f\u7528\uff09

## \u76ee\u9304\u7d50\u69cb

- `super_converter.py`
  - \u4e3b\u8981\u8f49\u6a94\u908f\u8f2f\u8207 GUI
- `launch_super_converter.pyw`
  - Windows GUI \u555f\u52d5\u9032\u5165\u9ede
- `build_portable.ps1`
  - \u7522\u751f portable \u7248\u672c
- `tools\\pandoc\\pandoc.exe`
  - portable \u6253\u5305\u7528 pandoc
- `portable\\WanyongConverter`
  - \u5df2\u6253\u5305\u5b8c\u7684 portable \u6210\u54c1

## \u57f7\u884c

```powershell
python .\super_converter.py
```

\u6216\u8005\uff1a

```powershell
python .\launch_super_converter.pyw
```

## \u6253\u5305 portable

```powershell
powershell -ExecutionPolicy Bypass -File .\build_portable.ps1
```

\u8f38\u51fa\u6703\u5728\uff1a`portable\\WanyongConverter\\`

## portable \u4f7f\u7528\u65b9\u5f0f

1. \u958b\u555f `portable\\WanyongConverter\\WanyongConverter.exe`
2. \u8981\u5206\u4eab\u7d66\u5225\u4eba\u6642\uff0c\u6574\u500b `WanyongConverter` \u8cc7\u6599\u593e\u4e00\u8d77\u5e36\u8d70
3. \u522a\u6389\u6574\u500b `portable\\WanyongConverter` \u8cc7\u6599\u593e\u5373\u53ef\u79fb\u9664

## GitHub \u5099\u8a3b

\u9019\u500b repo \u5305\u542b portable \u6210\u54c1\u8207\u5927\u578b\u57f7\u884c\u6a94\uff0c\u56e0\u6b64 `pandoc.exe` \u4f7f\u7528 Git LFS \u8ffd\u8e64\u3002
