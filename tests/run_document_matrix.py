import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image
from pypdf import PdfReader

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from super_converter import SuperConverter, can_automate_word, pythoncom, win32_client  # noqa: E402


def noop_log(_: str) -> None:
    return


def run_pandoc(converter: SuperConverter, source: Path, target: Path) -> None:
    pandoc_cmd = converter.tools.pandoc_cmd or "pandoc"
    subprocess.run([pandoc_cmd, str(source), "-o", str(target)], check=True)


def create_doc_with_word(docx_path: Path, doc_path: Path) -> None:
    if not can_automate_word():
        raise RuntimeError("Microsoft Word is not available")

    pythoncom.CoInitialize()
    word_app = None
    document = None
    try:
        word_app = win32_client.DispatchEx("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0
        document = word_app.Documents.Open(
            str(docx_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        document.SaveAs2(str(doc_path), FileFormat=0)
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word_app is not None:
            try:
                word_app.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def create_source_documents(converter: SuperConverter, temp_root: Path) -> dict[str, Path]:
    # 這份內容故意放中文、英文、清單與段落，避免只測到最單純的情況。
    base_markdown = """# 文件矩陣測試

這是一份文件轉檔測試。

This is a document conversion test.

- 項目一
- 項目二

## 小節

第二段內容。
"""

    plain_text = "文件矩陣測試\n\n這是一份 TXT 測試文件。\nThis is a TXT test document.\n"
    html_text = "<h1>文件矩陣測試</h1><p>這是一份 HTML 測試文件。</p><p>This is an HTML test document.</p>"
    rst_text = "文件矩陣測試\n============\n\n這是一份 RST 測試文件。\n\n- 項目一\n- 項目二\n"
    org_text = "* 文件矩陣測試\n\n這是一份 Org 測試文件。\n\n- 項目一\n- 項目二\n"
    tex_text = "\\section{文件矩陣測試}\n這是一份 TeX 測試文件。\n"
    docbook_text = """<?xml version="1.0" encoding="UTF-8"?>
<article xmlns="http://docbook.org/ns/docbook" version="5.0">
  <title>文件矩陣測試</title>
  <para>這是一份 DocBook 測試文件。</para>
</article>
"""

    sources: dict[str, Path] = {}

    def write_text_file(name: str, content: str) -> None:
        path = temp_root / name
        path.write_text(content, encoding="utf-8")
        sources[path.suffix.lower()] = path

    write_text_file("sample.txt", plain_text)
    write_text_file("sample.md", base_markdown)
    write_text_file("sample.markdown", base_markdown)
    write_text_file("sample.html", html_text)
    write_text_file("sample.htm", html_text)
    write_text_file("sample.rst", rst_text)
    write_text_file("sample.org", org_text)
    write_text_file("sample.tex", tex_text)
    write_text_file("sample.latex", tex_text)
    write_text_file("sample.docbook", docbook_text)

    markdown_source = sources[".md"]
    for ext in [".docx", ".odt", ".rtf", ".epub"]:
        target = temp_root / f"sample{ext}"
        run_pandoc(converter, markdown_source, target)
        sources[ext] = target

    pdf_target = temp_root / "sample.pdf"
    converter.markdown_to_pdf(base_markdown, pdf_target, base_dir=temp_root)
    sources[".pdf"] = pdf_target

    if can_automate_word():
        doc_target = temp_root / "sample.doc"
        create_doc_with_word(sources[".docx"], doc_target)
        sources[".doc"] = doc_target

    return sources


def verify_output(target_path: Path) -> None:
    ext = target_path.suffix.lower()
    if not target_path.exists() or target_path.stat().st_size == 0:
        raise ValueError("output file missing or empty")

    if ext == ".pdf":
        with target_path.open("rb") as handle:
            if handle.read(4) != b"%PDF":
                raise ValueError("output is not a valid PDF")
        return

    if ext in {".txt", ".md"}:
        text = target_path.read_text(encoding="utf-8", errors="ignore").strip()
        if not text:
            raise ValueError("text output is empty")
        return

    if ext == ".html":
        text = target_path.read_text(encoding="utf-8", errors="ignore")
        if "<" not in text:
            raise ValueError("html output does not look like HTML")
        return

    if ext == ".rtf":
        text = target_path.read_text(encoding="utf-8", errors="ignore")
        if "{\\" not in text or "\\par" not in text:
            raise ValueError("rtf output does not look like RTF content")
        return

    if ext == ".docx":
        with zipfile.ZipFile(target_path) as archive:
            if "word/document.xml" not in archive.namelist():
                raise ValueError("docx output is missing word/document.xml")
        return

    if ext == ".odt":
        with zipfile.ZipFile(target_path) as archive:
            if "content.xml" not in archive.namelist():
                raise ValueError("odt output is missing content.xml")
        return

    if ext == ".epub":
        with zipfile.ZipFile(target_path) as archive:
            names = set(archive.namelist())
            if "mimetype" not in names:
                raise ValueError("epub output is missing mimetype")
        return

    if ext in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".ico", ".apng", ".avif", ".avifs",
               ".bw", ".dds", ".dib", ".icb", ".icns", ".im", ".j2c", ".j2k", ".jfif", ".jp2", ".jpc", ".jpe",
               ".jpf", ".jpx", ".mpo", ".msp", ".pbm", ".pcx", ".pfm", ".pgm", ".pnm", ".ppm", ".qoi", ".rgb",
               ".rgba", ".sgi", ".tga", ".vda", ".vst", ".xbm"}:
        with Image.open(target_path) as img:
            img.load()
            if img.width <= 0 or img.height <= 0:
                raise ValueError("image output has invalid size")
        return

    raise ValueError(f"no verifier implemented for {ext}")


def main() -> None:
    converter = SuperConverter()
    results: list[dict[str, str]] = []

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        sources = create_source_documents(converter, temp_root)

        print("=== Document Source Files ===")
        for ext in sorted(sources):
            print(f"{ext}: {sources[ext].name}")
        print("")

        for source_ext, source_path in sorted(sources.items()):
            supported_outputs = converter.get_supported_outputs(source_path)
            for target_ext in supported_outputs:
                target_path = temp_root / f"{source_path.stem}_to_{target_ext.lstrip('.')}{target_ext}"
                result = {
                    "source_ext": source_ext,
                    "target_ext": target_ext,
                    "status": "success",
                    "error": "",
                }
                try:
                    converter.convert(source_path, target_path, noop_log)
                    verify_output(target_path)
                except Exception as exc:
                    result["status"] = "failed"
                    result["error"] = str(exc)
                results.append(result)

    success_items = [item for item in results if item["status"] == "success"]
    failed_items = [item for item in results if item["status"] == "failed"]

    print("=== Document Conversion Matrix Summary ===")
    print(f"Attempted: {len(results)}")
    print(f"Success:   {len(success_items)}")
    print(f"Failed:    {len(failed_items)}")
    print("")

    print("=== Successful Pairs ===")
    for item in success_items:
        print(f"{item['source_ext']} -> {item['target_ext']}")
    print("")

    print("=== Failed Pairs ===")
    for item in failed_items:
        print(f"{item['source_ext']} -> {item['target_ext']} | {item['error']}")


if __name__ == "__main__":
    main()
