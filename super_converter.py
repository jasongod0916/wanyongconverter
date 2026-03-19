import ctypes
import html as html_lib
import locale
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import scrolledtext
from tkinter import font as tkfont
from typing import Callable

import tkinter as tk
from markdown import markdown as md_to_html
from PIL import Image
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    import imageio_ffmpeg
except Exception:
    imageio_ffmpeg = None

try:
    from pdf2docx import Converter as PdfToDocxConverter
except Exception:
    PdfToDocxConverter = None

try:
    import pythoncom
    import win32com.client as win32_client
except Exception:
    pythoncom = None
    win32_client = None

TEXT_INPUTS = {'.txt', '.md', '.markdown', '.html', '.htm', '.rst', '.tex', '.latex', '.org'}
PANDOC_INPUTS = TEXT_INPUTS | {'.docx', '.odt', '.rtf', '.epub', '.docbook'}
WORD_INPUTS = {'.doc', '.docx'}
DOC_INPUTS = PANDOC_INPUTS | {'.pdf', '.doc'}
VIDEO_INPUTS = {
    '.mp4', '.mov', '.avi', '.mkv', '.webm', '.flv', '.wmv', '.m4v', '.mp3', '.wav',
    '.aac', '.flac', '.ogg', '.m4a', '.3gp', '.mpeg', '.mpg', '.ts', '.m2ts', '.vob',
    '.rmvb', '.asf', '.aiff', '.wma', '.opus', '.amr'
}
VIDEO_OUTPUTS = ['.mp4', '.mov', '.avi', '.mkv', '.webm', '.flv', '.wmv', '.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.gif']
PANDOC_OUTPUTS = {'.docx', '.odt', '.epub', '.rtf', '.html', '.txt', '.md'}
DOCUMENT_OUTPUTS = sorted(PANDOC_OUTPUTS | {'.pdf'})


@dataclass
class ToolAvailability:
    ffmpeg: bool
    pandoc: bool
    ffmpeg_cmd: str | None
    pandoc_cmd: str | None
    ffmpeg_source: str
    pandoc_source: str


def app_root() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_existing_path(candidates: list[Path]) -> str | None:
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def choose_ui_font(root: tk.Tk) -> str:
    preferred = ['Microsoft JhengHei UI', 'Microsoft JhengHei', '\u5fae\u8edf\u6b63\u9ed1\u9ad4', 'Noto Sans TC', 'Segoe UI']
    available = set(tkfont.families(root))
    for name in preferred:
        if name in available:
            return name
    return 'TkDefaultFont'


PDF_CJK_FONT_NAME = 'STSong-Light'
PDF_FONT_NAME = 'WanyongPdfRegular'
PDF_FONT_BOLD_NAME = 'WanyongPdfBold'


def register_pdf_fonts() -> tuple[str, str]:
    fonts_dir = Path(r'C:\Windows\Fonts')
    candidates = [
        ('msjh.ttc', 'msjhbd.ttc'),
        ('msyh.ttc', 'msyhbd.ttc'),
        ('mingliu.ttc', 'mingliub.ttc'),
        ('simsun.ttc', 'simsunb.ttf'),
        ('kaiu.ttf', 'kaiu.ttf'),
    ]
    for normal_file, bold_file in candidates:
        normal_path = fonts_dir / normal_file
        bold_path = fonts_dir / bold_file
        if not normal_path.exists() or not bold_path.exists():
            continue
        try:
            pdfmetrics.registerFont(TTFont(PDF_FONT_NAME, str(normal_path)))
            pdfmetrics.registerFont(TTFont(PDF_FONT_BOLD_NAME, str(bold_path)))
            pdfmetrics.registerFontFamily(PDF_FONT_NAME, normal=PDF_FONT_NAME, bold=PDF_FONT_BOLD_NAME)
            return PDF_FONT_NAME, PDF_FONT_BOLD_NAME
        except Exception:
            continue
    pdfmetrics.registerFont(UnicodeCIDFont(PDF_CJK_FONT_NAME))
    return PDF_CJK_FONT_NAME, PDF_CJK_FONT_NAME


PDF_FONT_NAME, PDF_FONT_BOLD_NAME = register_pdf_fonts()


_WORD_AUTOMATION_AVAILABLE: bool | None = None


def can_automate_word() -> bool:
    global _WORD_AUTOMATION_AVAILABLE
    if _WORD_AUTOMATION_AVAILABLE is not None:
        return _WORD_AUTOMATION_AVAILABLE
    if pythoncom is None or win32_client is None:
        _WORD_AUTOMATION_AVAILABLE = False
        return _WORD_AUTOMATION_AVAILABLE
    initialized = False
    word_app = None
    try:
        pythoncom.CoInitialize()
        initialized = True
        word_app = win32_client.DispatchEx('Word.Application')
        word_app.Visible = False
        word_app.DisplayAlerts = 0
        _WORD_AUTOMATION_AVAILABLE = True
    except Exception:
        _WORD_AUTOMATION_AVAILABLE = False
    finally:
        if word_app is not None:
            try:
                word_app.Quit(False)
            except Exception:
                pass
        if initialized:
            pythoncom.CoUninitialize()
    return _WORD_AUTOMATION_AVAILABLE


def get_default_output_dir() -> Path:
    for candidate in [Path.home() / 'Desktop', Path.home() / 'OneDrive' / 'Desktop', Path.home()]:
        if candidate.exists():
            return candidate
    return Path.cwd()
def is_ascii_path(path: Path | str) -> bool:
    try:
        str(path).encode('ascii')
        return True
    except UnicodeEncodeError:
        return False


def get_ascii_runtime_dir() -> Path:
    runtime_dir = Path.home() / 'AppData' / 'Local' / 'WanyongConverter' / 'runtime'
    runtime_dir.mkdir(parents=True, exist_ok=True)
    return runtime_dir


def write_startup_trace(message: str) -> None:
    desktop = Path.home() / 'Desktop'
    if not desktop.exists():
        desktop = Path.home()
    trace_path = desktop / 'wanyong_converter_startup_trace.txt'
    with trace_path.open('a', encoding='utf-8') as handle:
        handle.write(message + '\n')


def detect_tools() -> ToolAvailability:
    tools_dir = app_root() / 'tools'
    local_ffmpeg = find_existing_path([tools_dir / 'ffmpeg' / 'ffmpeg.exe', tools_dir / 'ffmpeg.exe'])
    local_pandoc = find_existing_path([tools_dir / 'pandoc' / 'pandoc.exe', tools_dir / 'pandoc.exe'])
    ffmpeg_cmd = local_ffmpeg
    ffmpeg_source = 'portable' if local_ffmpeg else 'missing'
    if not ffmpeg_cmd and imageio_ffmpeg is not None:
        try:
            ffmpeg_cmd = imageio_ffmpeg.get_ffmpeg_exe()
            ffmpeg_source = 'embedded'
        except Exception:
            ffmpeg_cmd = None
    return ToolAvailability(
        ffmpeg=ffmpeg_cmd is not None,
        pandoc=local_pandoc is not None,
        ffmpeg_cmd=ffmpeg_cmd,
        pandoc_cmd=local_pandoc,
        ffmpeg_source=ffmpeg_source,
        pandoc_source='portable' if local_pandoc else 'missing',
    )


class ConversionError(RuntimeError):
    pass


class SuperConverter:
    def __init__(self) -> None:
        self.tools = detect_tools()
        self.word_available: bool | None = None

    def get_word_availability(self) -> bool:
        if self.word_available is None:
            self.word_available = can_automate_word()
        return self.word_available

    def get_supported_outputs(self, input_path: Path) -> list[str]:
        suffix = input_path.suffix.lower()
        if self.is_image_file(input_path):
            return [ext for ext in self.get_image_outputs() if ext != suffix]
        if suffix in DOC_INPUTS:
            return [ext for ext in DOCUMENT_OUTPUTS if ext != suffix and self.can_convert_document(input_path, ext)]
        if suffix in VIDEO_INPUTS:
            return [ext for ext in VIDEO_OUTPUTS if ext != suffix] if self.tools.ffmpeg else []
        return []

    def can_convert_document(self, src: Path, target_ext: str) -> bool:
        src_ext = src.suffix.lower()
        if src_ext in TEXT_INPUTS and target_ext == '.pdf':
            return True
        if src_ext == '.docx' and target_ext == '.pdf':
            return self.tools.pandoc or (pythoncom is not None and win32_client is not None)
        if src_ext == '.doc' and target_ext == '.pdf':
            return pythoncom is not None and win32_client is not None
        if src_ext in PANDOC_INPUTS and target_ext == '.pdf':
            return self.tools.pandoc
        if src_ext == '.pdf' and target_ext == '.docx':
            return PdfToDocxConverter is not None
        if src_ext == '.pdf' and target_ext in {'.txt', '.md'}:
            return True
        if src_ext in PANDOC_INPUTS and target_ext in PANDOC_OUTPUTS:
            return self.tools.pandoc
        return False

    def get_image_inputs(self) -> list[str]:
        common = {'.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif', '.tif', '.tiff', '.ico'}
        extensions = {ext.lower() for ext in Image.registered_extensions()}
        extensions.discard('.pdf')
        return sorted(common | extensions)

    def get_image_outputs(self) -> list[str]:
        return sorted(set(self.get_image_inputs()) | {'.pdf'})

    def is_image_file(self, source: Path) -> bool:
        return source.suffix.lower() in self.get_image_inputs()

    def convert(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        suffix = source.suffix.lower()
        log(f'\u958b\u59cb\u8655\u7406\u6a94\u6848: {source.name}')
        log(f'\u4f86\u6e90\u8def\u5f91: {source}')
        log(f'\u8f38\u51fa\u76ee\u6a19: {target}')
        if self.is_image_file(source):
            log('\u8f49\u6a94\u985e\u578b: \u5716\u7247')
            self.convert_image(source, target, log)
            return
        if suffix in DOC_INPUTS:
            log('\u8f49\u6a94\u985e\u578b: \u6587\u4ef6')
            self.convert_document(source, target, log)
            return
        if suffix in VIDEO_INPUTS:
            log('\u8f49\u6a94\u985e\u578b: \u5f71\u97f3')
            self.convert_media(source, target, log)
            return
        raise ConversionError(f'\u4e0d\u652f\u63f4\u7684\u8f38\u5165\u683c\u5f0f: {suffix}')

    def convert_image(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        target_ext = target.suffix.lower()
        log(f'\u5716\u7247\u8f38\u51fa\u683c\u5f0f: {target_ext}')
        with Image.open(source) as img:
            if target_ext in {'.jpg', '.jpeg'}:
                if img.mode in ('RGBA', 'LA', 'P'):
                    img = img.convert('RGB')
                img.save(target, quality=95)
            elif target_ext == '.pdf':
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                img.save(target, 'PDF', resolution=100.0)
            else:
                if img.mode == 'RGBA' and target_ext in {'.bmp', '.jpg', '.jpeg'}:
                    img = img.convert('RGB')
                img.save(target)
        log('\u5716\u7247\u8f49\u6a94\u5b8c\u6210')

    def convert_document(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        src_ext = source.suffix.lower()
        dst_ext = target.suffix.lower()
        log(f'\u6587\u4ef6\u8f49\u6a94\u8def\u5f91: {src_ext} -> {dst_ext}')
        if src_ext in TEXT_INPUTS and dst_ext == '.pdf':
            self.convert_text_like(source, target, log)
            return
        if src_ext in WORD_INPUTS and dst_ext == '.pdf':
            log('\u50c5\u5728 DOC/DOCX -> PDF \u958b\u59cb\u8f49\u6a94\u6642\u6aa2\u67e5 Microsoft Word')
            if self.get_word_availability():
                self.convert_word_to_pdf(source, target, log)
                return
            if src_ext in PANDOC_INPUTS:
                log('\u672a\u5075\u6e2c\u5230 Microsoft Word\uff0c\u6539\u7528 portable PDF \u8def\u7dda')
                self.convert_pandoc_to_pdf(source, target, log)
                return
            raise ConversionError('\u672a\u5075\u6e2c\u5230 Microsoft Word\uff0cDOC \u8f49 PDF \u9700\u8981 Word')
        if src_ext in PANDOC_INPUTS and dst_ext == '.pdf':
            self.convert_pandoc_to_pdf(source, target, log)
            return
        if src_ext == '.pdf' and dst_ext == '.docx':
            self.convert_pdf_to_docx(source, target, log)
            return
        if src_ext == '.pdf' and dst_ext == '.md':
            self.convert_pdf_to_markdown(source, target, log)
            return
        if src_ext == '.pdf' and dst_ext == '.txt':
            self.convert_pdf_to_text(source, target, log)
            return
        if src_ext in PANDOC_INPUTS and dst_ext in PANDOC_OUTPUTS:
            self.convert_with_pandoc(source, target, log)
            return
        raise ConversionError(f'\u76ee\u524d\u7121\u6cd5\u8f49\u63db\u9019\u7d44\u6587\u4ef6\u683c\u5f0f: {src_ext} -> {dst_ext}')

    def convert_text_like(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        src_ext = source.suffix.lower()
        dst_ext = target.suffix.lower()
        log(f'\u6587\u5b57\u8655\u7406\u6a21\u5f0f: {src_ext} -> {dst_ext}')
        text = source.read_text(encoding='utf-8', errors='ignore')
        if dst_ext in {'.txt', '.md'}:
            target.write_text(text, encoding='utf-8')
            log('\u6587\u5b57\u6216 Markdown \u8f49\u6a94\u5b8c\u6210')
            return
        if dst_ext == '.html':
            if src_ext in {'.md', '.markdown'}:
                html = md_to_html(text)
            elif src_ext == '.txt':
                html = '<pre>' + self.escape_html(text) + '</pre>'
            else:
                html = text
            target.write_text(html, encoding='utf-8')
            log('HTML \u8f49\u6a94\u5b8c\u6210')
            return
        if dst_ext == '.pdf':
            if src_ext in {'.md', '.markdown'}:
                self.markdown_to_pdf(text, target, base_dir=source.parent)
            elif src_ext in {'.html', '.htm'}:
                self.html_to_pdf(text, target)
            else:
                self.text_to_pdf(text, target)
            log('PDF \u8f49\u6a94\u5b8c\u6210')
            return
        if dst_ext in PANDOC_OUTPUTS:
            if not self.tools.pandoc:
                raise ConversionError('\u9700\u8981 pandoc \u624d\u80fd\u5b8c\u6210\u9019\u500b\u6587\u4ef6\u8f49\u63db')
            self.convert_with_pandoc(source, target, log)
            return
        raise ConversionError(f'\u4e0d\u652f\u63f4\u7684\u6587\u5b57\u8f38\u51fa\u683c\u5f0f: {dst_ext}')

    def run_with_ascii_paths(self, source: Path, target: Path, log: Callable[[str], None], worker: Callable[[Path, Path], None]) -> None:
        runtime_dir = get_ascii_runtime_dir()
        with tempfile.TemporaryDirectory(dir=str(runtime_dir)) as temp_dir:
            temp_root = Path(temp_dir)
            staged_source = temp_root / f'source{source.suffix.lower()}'
            staged_target = temp_root / f'output{target.suffix.lower()}'
            shutil.copy2(source, staged_source)
            worker(staged_source, staged_target)
            if not staged_target.exists():
                raise ConversionError('\u5de5\u5177\u57f7\u884c\u5b8c\u6210\uff0c\u4f46\u6c92\u6709\u7522\u51fa\u8f38\u51fa\u6a94\u6848')
            shutil.copy2(staged_target, target)

    def convert_word_to_pdf(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        if not self.get_word_availability():
            raise ConversionError('\u76ee\u524d\u7121\u6cd5\u547c\u53eb Microsoft Word \u9032\u884c PDF \u532f\u51fa')
        log('\u5075\u6e2c\u5230 Microsoft Word\uff0c\u512a\u5148\u4f7f\u7528\u9ad8\u9084\u539f PDF \u532f\u51fa')

        def worker(staged_source: Path, staged_target: Path) -> None:
            initialized = False
            word_app = None
            document = None
            try:
                pythoncom.CoInitialize()
                initialized = True
                word_app = win32_client.DispatchEx('Word.Application')
                word_app.Visible = False
                word_app.DisplayAlerts = 0
                log('\u4f7f\u7528 Microsoft Word \u81ea\u52d5\u8f38\u51fa PDF')
                document = word_app.Documents.Open(
                    str(staged_source),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                )
                document.ExportAsFixedFormat(str(staged_target), 17)
            except Exception as exc:
                raise ConversionError(f'Microsoft Word PDF \u532f\u51fa\u5931\u6557: {exc}') from exc
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
                if initialized:
                    pythoncom.CoUninitialize()

        self.run_with_ascii_paths(source, target, log, worker)
        log('Word PDF \u532f\u51fa\u5b8c\u6210')

    def convert_pandoc_to_pdf(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        log('\u5148\u7528 pandoc \u8f49\u6210\u66ab\u5b58 Markdown\uff0c\u518d\u4ee5 PDF \u7248\u578b\u8f38\u51fa')
        runtime_dir = get_ascii_runtime_dir()
        with tempfile.TemporaryDirectory(dir=str(runtime_dir)) as temp_dir:
            temp_root = Path(temp_dir)
            staged_source = temp_root / f'source{source.suffix.lower()}'
            temp_md = temp_root / 'intermediate.md'
            shutil.copy2(source, staged_source)
            log('\u4f7f\u7528 pandoc \u6574\u7406 DOCX \u7d50\u69cb\u70ba\u8f03\u7a69\u5b9a\u7684 Markdown \u4e2d\u4ecb')
            self.run_command([
                self.tools.pandoc_cmd or 'pandoc',
                str(staged_source),
                '--to=gfm+pipe_tables',
                '--wrap=none',
                '--extract-media=' + str(temp_root),
                '-o',
                str(temp_md),
            ], log)
            markdown = temp_md.read_text(encoding='utf-8', errors='ignore')
            self.markdown_to_pdf(markdown, target, base_dir=temp_root)
        log('Pandoc PDF \u8f49\u6a94\u5b8c\u6210')

    def convert_media(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        if not self.tools.ffmpeg:
            raise ConversionError('\u9700\u8981 ffmpeg \u624d\u80fd\u9032\u884c\u5f71\u97f3\u8f49\u6a94')
        log('\u4f7f\u7528 ffmpeg \u9032\u884c\u5f71\u97f3\u8f49\u6a94')
        cmd = [self.tools.ffmpeg_cmd or 'ffmpeg', '-y', '-i', str(source)]
        dst_ext = target.suffix.lower()
        if dst_ext == '.gif':
            cmd += ['-vf', 'fps=12,scale=720:-1:flags=lanczos']
        elif dst_ext == '.mp3':
            cmd += ['-vn', '-c:a', 'libmp3lame', '-q:a', '2']
        elif dst_ext == '.wav':
            cmd += ['-vn', '-c:a', 'pcm_s16le']
        elif dst_ext == '.mp4':
            cmd += ['-c:v', 'libx264', '-c:a', 'aac']
        elif dst_ext == '.webm':
            cmd += ['-c:v', 'libvpx-vp9', '-c:a', 'libopus']
        elif dst_ext == '.mov':
            cmd += ['-c:v', 'mpeg4', '-c:a', 'aac']
        elif dst_ext == '.avi':
            cmd += ['-c:v', 'mpeg4', '-c:a', 'mp3']
        elif dst_ext == '.mkv':
            cmd += ['-c:v', 'libx264', '-c:a', 'aac']
        elif dst_ext == '.flv':
            cmd += ['-c:v', 'flv', '-c:a', 'libmp3lame']
        elif dst_ext == '.wmv':
            cmd += ['-c:v', 'wmv2', '-c:a', 'wmav2']
        elif dst_ext == '.aac':
            cmd += ['-vn', '-c:a', 'aac', '-b:a', '192k']
        elif dst_ext == '.flac':
            cmd += ['-vn', '-c:a', 'flac']
        elif dst_ext == '.ogg':
            cmd += ['-vn', '-c:a', 'libvorbis', '-q:a', '5']
        elif dst_ext == '.m4a':
            cmd += ['-vn', '-c:a', 'aac', '-b:a', '192k']
        cmd.append(str(target))
        self.run_command(cmd, log)
        log('\u5f71\u97f3\u8f49\u6a94\u5b8c\u6210')

    def convert_pdf_to_docx(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        if PdfToDocxConverter is None:
            raise ConversionError('\u9700\u8981 pdf2docx \u624d\u80fd\u628a PDF \u8f49\u6210 DOCX')
        log('\u4f7f\u7528 pdf2docx \u5c07 PDF \u8f49\u6210 DOCX')

        def worker(staged_source: Path, staged_target: Path) -> None:
            converter = PdfToDocxConverter(str(staged_source))
            try:
                converter.convert(str(staged_target), start=0, end=None)
            finally:
                converter.close()

        self.run_with_ascii_paths(source, target, log, worker)
        log('PDF \u8f49 DOCX \u5b8c\u6210')

    def convert_pdf_to_markdown(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        log('\u5f9e PDF \u64f7\u53d6\u6587\u5b57\u4e26\u8f38\u51fa\u70ba Markdown')
        with source.open('rb') as handle:
            reader = PdfReader(handle)
            content = "\n\n".join(page.extract_text() or '' for page in reader.pages)
        target.write_text(content, encoding='utf-8')
        log('PDF \u8f49 Markdown \u5b8c\u6210')

    def convert_pdf_to_text(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        log('\u5f9e PDF \u64f7\u53d6\u6587\u5b57\u4e26\u8f38\u51fa\u70ba TXT')
        with source.open('rb') as handle:
            reader = PdfReader(handle)
            content = "\n\n".join(page.extract_text() or '' for page in reader.pages)
        target.write_text(content, encoding='utf-8')
        log('PDF \u8f49 TXT \u5b8c\u6210')

    def convert_with_pandoc(self, source: Path, target: Path, log: Callable[[str], None]) -> None:
        if not self.tools.pandoc:
            raise ConversionError('\u9700\u8981 pandoc \u624d\u80fd\u9032\u884c\u9019\u500b\u6587\u4ef6\u8f49\u63db')
        log('\u4f7f\u7528 pandoc \u9032\u884c\u6587\u4ef6\u8f49\u6a94')
        if is_ascii_path(source) and is_ascii_path(target):
            self.run_command([self.tools.pandoc_cmd or 'pandoc', str(source), '-o', str(target)], log)
        else:
            self.run_with_ascii_paths(
                source,
                target,
                log,
                lambda staged_source, staged_target: self.run_command(
                    [self.tools.pandoc_cmd or 'pandoc', str(staged_source), '-o', str(staged_target)],
                    log,
                ),
            )
        log('Pandoc \u8f49\u6a94\u5b8c\u6210')

    def run_command(self, cmd: list[str], log: Callable[[str], None]) -> None:
        encoding = locale.getpreferredencoding(False) or 'utf-8'
        log('\u57f7\u884c\u547d\u4ee4: ' + ' '.join(cmd))
        result = subprocess.run(cmd, capture_output=True, text=True, encoding=encoding, errors='replace', check=False)
        if result.stdout.strip():
            for line in result.stdout.strip().splitlines():
                log('stdout | ' + line)
        if result.stderr.strip():
            for line in result.stderr.strip().splitlines():
                log('stderr | ' + line)
        if result.returncode != 0:
            raise ConversionError(result.stderr.strip() or '\u547d\u4ee4\u57f7\u884c\u5931\u6557')
        log(f'\u547d\u4ee4\u5b8c\u6210\uff0c\u56de\u50b3\u78bc: {result.returncode}')

    def build_pdf_styles(self) -> dict[str, ParagraphStyle]:
        styles = getSampleStyleSheet()
        body = styles['BodyText'].clone('CJKBodyText')
        body.fontName = PDF_FONT_NAME
        body.fontSize = 10.5
        body.leading = 17
        body.spaceAfter = 7
        body.wordWrap = 'CJK'

        heading1 = ParagraphStyle(
            'CJKHeading1',
            parent=body,
            fontName=PDF_FONT_BOLD_NAME,
            fontSize=18,
            leading=24,
            spaceBefore=10,
            spaceAfter=10,
        )
        heading2 = ParagraphStyle(
            'CJKHeading2',
            parent=body,
            fontName=PDF_FONT_BOLD_NAME,
            fontSize=14,
            leading=20,
            spaceBefore=8,
            spaceAfter=8,
        )
        heading3 = ParagraphStyle(
            'CJKHeading3',
            parent=body,
            fontName=PDF_FONT_BOLD_NAME,
            fontSize=12,
            leading=18,
            spaceBefore=5,
            spaceAfter=6,
        )
        bullet = ParagraphStyle(
            'CJKBullet',
            parent=body,
            leftIndent=16,
            firstLineIndent=-12,
            spaceAfter=4,
        )
        return {
            'body': body,
            'heading1': heading1,
            'heading2': heading2,
            'heading3': heading3,
            'bullet': bullet,
        }

    def normalize_pdf_text(self, text: str) -> str:
        cleaned = text.replace('\r\n', '\n').replace('\r', '\n').replace('\f', '\n')
        replacements = {
            '\u00a0': ' ',
            '\u3000': ' ',
            '\u2013': '-',
            '\u2014': '-',
            '\u2022': '- ',
            '\u25cf': '- ',
            '\u25aa': '- ',
            '\u25ab': '- ',
            '\u2018': "'",
            '\u2019': "'",
            '\u201c': '"',
            '\u201d': '"',
            '\u2192': '->',
            '\u2026': '...',
        }
        for old, new in replacements.items():
            cleaned = cleaned.replace(old, new)
        cleaned = ''.join(ch for ch in cleaned if ch == '\n' or ch == '\t' or ord(ch) >= 32)
        cleaned = re.sub(r'[ \t]+\n', '\n', cleaned)
        cleaned = re.sub(r'\n[ \t]+', '\n', cleaned)
        cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
        return cleaned.strip()

    def simplify_inline_markdown(self, text: str) -> str:
        simplified = text
        simplified = re.sub(r'<!--.*?-->', '', simplified, flags=re.S)
        simplified = re.sub(r'</?span[^>]*>', '', simplified)
        simplified = re.sub(r'</?a[^>]*>', '', simplified)
        simplified = re.sub(r'!\[([^\]]*)\]\([^\)]+\)', r'\1', simplified)
        simplified = re.sub(r'\[([^\]]+)\]\([^\)]+\)', r'\1', simplified)
        simplified = re.sub(r'`([^`]+)`', r'\1', simplified)
        simplified = re.sub(r'\\([#*_`\\\[\](){}.!+-])', r'\1', simplified)
        return simplified.strip()

    def strip_inline_formatting(self, text: str) -> str:
        plain = self.simplify_inline_markdown(text)
        plain = re.sub(r'\*\*([^*]+)\*\*', r'\1', plain)
        plain = re.sub(r'__([^_]+)__', r'\1', plain)
        plain = re.sub(r'\*([^*]+)\*', r'\1', plain)
        plain = re.sub(r'_([^_]+)_', r'\1', plain)
        plain = re.sub(r'~~([^~]+)~~', r'\1', plain)
        plain = re.sub(r'</?u>', '', plain)
        return plain.strip()

    def format_pdf_inline_markup(self, text: str) -> str:
        formatted = self.simplify_inline_markdown(self.normalize_pdf_text(text))
        formatted = formatted.replace('<u>', 'ZZZUOPENTAGZZZ').replace('</u>', 'ZZZUCLOSETAGZZZ')
        formatted = html_lib.escape(formatted, quote=False)
        formatted = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', formatted)
        formatted = re.sub(r'__([^_]+)__', r'<b>\1</b>', formatted)
        formatted = re.sub(r'(?<!\*)\*([^*\n]+)\*(?!\*)', r'<i>\1</i>', formatted)
        formatted = re.sub(r'(?<!_)_([^_\n]+)_(?!_)', r'<i>\1</i>', formatted)
        formatted = formatted.replace('ZZZUOPENTAGZZZ', '<u>').replace('ZZZUCLOSETAGZZZ', '</u>')
        return formatted.strip()
    def clean_pandoc_markdown(self, text: str) -> str:
        cleaned = self.normalize_pdf_text(text)
        cleaned = re.sub(r'\{#[^}]+\}', '', cleaned)
        cleaned = re.sub(r'\[([^\]]+)\]\{[^}\n]+\}', r'\1', cleaned)
        cleaned = re.sub(r'\{[.#][^}\n]+\}', '', cleaned)
        cleaned = re.sub(r'\[\^\d+\]:.*', '', cleaned)
        cleaned = re.sub(r'(?m)^\s*`{3,}.*$', '', cleaned)
        cleaned = re.sub(r'(?m)^\s*:{3,}.*$', '', cleaned)
        cleaned = re.sub(r'(?m)^\s*[-*_]{3,}\s*$', '', cleaned)
        cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
        return cleaned.strip()

    def parse_length_to_points(self, value: str | None) -> float | None:
        if not value:
            return None
        match = re.match(r'([\d.]+)\s*(pt|px|in|cm|mm)?', value.strip(), flags=re.I)
        if not match:
            return None
        amount = float(match.group(1))
        unit = (match.group(2) or 'px').lower()
        factors = {
            'pt': 1.0,
            'px': 0.75,
            'in': inch,
            'cm': 72.0 / 2.54,
            'mm': 72.0 / 25.4,
        }
        return amount * factors.get(unit, 0.75)

    def parse_markdown_image(self, line: str) -> tuple[str, float | None, float | None] | None:
        html_match = re.search(r'<img\s+[^>]*src="([^"]+)"[^>]*?(?:style="([^"]*)")?[^>]*/?>', line, flags=re.I)
        if html_match:
            src = html_match.group(1).strip()
            style = html_match.group(2) or ''
            width = None
            height = None
            width_match = re.search(r'width:([\d.]+(?:pt|px|in|cm|mm)?)', style, flags=re.I)
            height_match = re.search(r'height:([\d.]+(?:pt|px|in|cm|mm)?)', style, flags=re.I)
            if width_match:
                width = self.parse_length_to_points(width_match.group(1))
            if height_match:
                height = self.parse_length_to_points(height_match.group(1))
            return src, width, height
        markdown_match = re.match(r'!\[[^\]]*\]\(([^)]+)\)', line)
        if markdown_match:
            return markdown_match.group(1).strip(), None, None
        return None

    def append_image_block(self, story: list, line: str, base_dir: Path | None) -> bool:
        parsed = self.parse_markdown_image(line)
        if not parsed:
            return False
        src, width, height = parsed
        image_path = Path(src)
        if not image_path.is_absolute() and base_dir is not None:
            image_path = (base_dir / image_path).resolve()
        if not image_path.exists():
            return False
        max_width = 170 * mm
        with Image.open(image_path) as img:
            img_width, img_height = img.size
        if width is None and height is None:
            width = min(max_width, img_width * 0.75)
            height = width * img_height / max(img_width, 1)
        elif width is None:
            width = height * img_width / max(img_height, 1)
        elif height is None:
            height = width * img_height / max(img_width, 1)
        if width and width > max_width:
            scale = max_width / width
            width *= scale
            height *= scale
        story.append(RLImage(str(image_path), width=width, height=height))
        story.append(Spacer(1, 8))
        return True

    def is_markdown_table_start(self, lines: list[str], index: int) -> bool:
        if index + 1 >= len(lines):
            return False
        header = lines[index].strip()
        divider = lines[index + 1].strip()
        if '|' not in header or header.count('|') < 2:
            return False
        return re.match(r'^\|?(?:\s*:?-+:?\s*\|)+\s*:?-+:?\s*\|?$', divider) is not None

    def parse_markdown_table(self, lines: list[str], start_index: int) -> tuple[list[list[str]], int]:
        rows: list[list[str]] = []
        index = start_index
        while index < len(lines):
            raw = lines[index].strip()
            if not raw:
                break
            if index == start_index + 1 and re.match(r'^\|?(?:\s*:?-+:?\s*\|)+\s*:?-+:?\s*\|?$', raw):
                index += 1
                continue
            if '|' not in raw or raw.count('|') < 2:
                break
            cells = [self.strip_inline_formatting(col.strip()) for col in raw.strip('|').split('|')]
            rows.append(cells)
            index += 1
        column_count = max((len(row) for row in rows), default=0)
        normalized = [row + [''] * (column_count - len(row)) for row in rows]
        return normalized, index

    def append_table_block(self, story: list, rows: list[list[str]], styles: dict[str, ParagraphStyle]) -> None:
        if not rows:
            return
        column_count = max(len(row) for row in rows)
        max_width = 170 * mm
        col_widths = [max_width / max(column_count, 1)] * column_count
        table_rows = []
        for row_index, row in enumerate(rows):
            style = styles['heading3'] if row_index == 0 else styles['body']
            table_rows.append([Paragraph(self.format_pdf_inline_markup(cell or ' '), style) for cell in row])
        table = Table(table_rows, colWidths=col_widths, repeatRows=1, hAlign='LEFT')
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dbe7ff')),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.HexColor('#10233f')),
            ('GRID', (0, 0), (-1, -1), 0.6, colors.HexColor('#8ea5c7')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(table)
        story.append(Spacer(1, 8))

    def append_plain_paragraph(self, story: list, text: str, style: ParagraphStyle) -> None:
        paragraph = self.format_pdf_inline_markup(text)
        if not paragraph:
            return
        paragraph = paragraph.replace('\n', '<br/>')
        story.append(Paragraph(paragraph, style))
        story.append(Spacer(1, 4))

    def markdown_to_story(self, markdown_text: str, base_dir: Path | None = None) -> list:
        styles = self.build_pdf_styles()
        story = []
        lines = self.clean_pandoc_markdown(markdown_text).split('\n')
        buffer: list[str] = []

        def flush_buffer() -> None:
            if buffer:
                merged = ' '.join(part.strip() for part in buffer if part.strip())
                self.append_plain_paragraph(story, merged, styles['body'])
                buffer.clear()

        index = 0
        while index < len(lines):
            raw_line = lines[index]
            line = self.simplify_inline_markdown(raw_line.strip())
            plain_line = self.strip_inline_formatting(line)
            if not plain_line:
                flush_buffer()
                index += 1
                continue
            if self.is_markdown_table_start(lines, index):
                flush_buffer()
                rows, index = self.parse_markdown_table(lines, index)
                self.append_table_block(story, rows, styles)
                continue
            if self.append_image_block(story, line, base_dir):
                flush_buffer()
                index += 1
                continue
            if line.startswith('# '):
                flush_buffer()
                story.append(Paragraph(self.format_pdf_inline_markup(line[2:].strip()), styles['heading1']))
                story.append(Spacer(1, 4))
                index += 1
                continue
            if line.startswith('## '):
                flush_buffer()
                story.append(Paragraph(self.format_pdf_inline_markup(line[3:].strip()), styles['heading2']))
                story.append(Spacer(1, 4))
                index += 1
                continue
            if line.startswith('### '):
                flush_buffer()
                story.append(Paragraph(self.format_pdf_inline_markup(line[4:].strip()), styles['heading3']))
                story.append(Spacer(1, 4))
                index += 1
                continue
            bold_heading = re.fullmatch(r'\*\*(.+?)\*\*', line)
            if bold_heading and len(self.strip_inline_formatting(bold_heading.group(1))) <= 80:
                flush_buffer()
                heading_text = bold_heading.group(1).strip()
                heading_style = styles['heading1'] if len(story) < 2 else styles['heading2']
                story.append(Paragraph(self.format_pdf_inline_markup(heading_text), heading_style))
                story.append(Spacer(1, 4))
                index += 1
                continue
            if line.startswith('>'):
                quote_text = re.sub(r'^>+\s*', '', line)
                buffer.append(quote_text)
                index += 1
                continue
            if re.match(r'^[-*]\s+', line):
                flush_buffer()
                bullet_text = re.sub(r'^[-*]\s+', '', line)
                story.append(Paragraph(self.format_pdf_inline_markup('- ' + bullet_text), styles['bullet']))
                story.append(Spacer(1, 2))
                index += 1
                continue
            if re.match(r'^\d+\.\s+', line):
                flush_buffer()
                story.append(Paragraph(self.format_pdf_inline_markup(line), styles['bullet']))
                story.append(Spacer(1, 2))
                index += 1
                continue
            buffer.append(line)
            index += 1

        flush_buffer()
        return story or [Paragraph('(empty)', styles['body'])]

    def markdown_to_pdf(self, markdown_text: str, target: Path, base_dir: Path | None = None) -> None:
        doc = SimpleDocTemplate(
            str(target),
            pagesize=A4,
            leftMargin=18 * mm,
            rightMargin=18 * mm,
            topMargin=16 * mm,
            bottomMargin=16 * mm,
        )
        doc.build(self.markdown_to_story(markdown_text, base_dir=base_dir))

    def text_to_pdf(self, text: str, target: Path) -> None:
        styles = self.build_pdf_styles()
        doc = SimpleDocTemplate(
            str(target),
            pagesize=A4,
            leftMargin=18 * mm,
            rightMargin=18 * mm,
            topMargin=16 * mm,
            bottomMargin=16 * mm,
        )
        story = []
        for block in self.normalize_pdf_text(text).split('\n\n'):
            self.append_plain_paragraph(story, block, styles['body'])
        doc.build(story or [Paragraph('(empty)', styles['body'])])

    def html_to_pdf(self, html: str, target: Path) -> None:
        body_text = '\n\n'.join(self.coerce_html_paragraphs(html))
        self.text_to_pdf(body_text, target)

    @staticmethod
    def coerce_html_paragraphs(html: str) -> list[str]:
        cleaned = html.replace('\r', '')
        cleaned = re.sub(r'(?is)<(script|style).*?>.*?</\1>', '', cleaned)
        cleaned = cleaned.replace('</p>', '\n\n').replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
        cleaned = re.sub(r'<[^>]+>', '', cleaned)
        cleaned = html_lib.unescape(cleaned)
        segments = []
        for chunk in cleaned.split('\n\n'):
            chunk = re.sub(r'\s+', ' ', chunk).strip()
            if chunk:
                segments.append(chunk)
        return segments or ['(empty)']

    @staticmethod
    def escape_html(text: str) -> str:
        return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.ui_font = choose_ui_font(root)
        self.mono_font = 'Consolas'
        self.root.title('\u842c\u7528\u8f49\u6a94\u738b')
        self.root.geometry('980x720')
        self.root.minsize(920, 680)
        self.converter = SuperConverter()
        self.selected_files: list[Path] = []
        self.file_output_map: dict[Path, list[str]] = {}
        self.progress_text_var = tk.StringVar(value='0 / 0')
        self.progress_value = tk.DoubleVar(value=0.0)
        self.output_dir_var = tk.StringVar(value=str(get_default_output_dir()))
        self.format_var = tk.StringVar()
        self.status_var = tk.StringVar(value='\u8acb\u5148\u9078\u64c7\u6a94\u6848\uff0c\u518d\u958b\u59cb\u8f49\u6a94')
        self.drop_hint_var = tk.StringVar(value='\u53ef\u76f4\u63a5\u628a\u6a94\u6848\u62d6\u9032\u4e0b\u65b9\u6e05\u55ae')
        self.configure_styles()
        self.build_ui()
        self.enable_file_drop()
        Path(self.output_dir_var.get()).mkdir(parents=True, exist_ok=True)

    def configure_styles(self) -> None:
        style = ttk.Style()
        if 'vista' in style.theme_names():
            style.theme_use('vista')
        self.root.configure(bg='#f3f6fb')
        style.configure('Shell.TFrame', background='#f3f6fb')
        style.configure('Card.TFrame', background='#ffffff')
        style.configure('HeroTitle.TLabel', background='#f3f6fb', foreground='#10233f', font=(self.ui_font, 24, 'bold'))
        style.configure('HeroSub.TLabel', background='#f3f6fb', foreground='#51627a', font=(self.ui_font, 11))
        style.configure('Section.TLabelframe', background='#ffffff')
        style.configure('Section.TLabelframe.Label', background='#ffffff', foreground='#19314f', font=(self.ui_font, 10, 'bold'))
        style.configure('Info.TLabel', background='#ffffff', foreground='#314760', font=(self.ui_font, 10))
        style.configure('Metric.TLabel', background='#ffffff', foreground='#0d223d', font=(self.ui_font, 10, 'bold'))
        style.configure('Muted.TLabel', background='#ffffff', foreground='#6d7d92', font=(self.ui_font, 9))
        style.configure('Accent.TButton', font=(self.ui_font, 10, 'bold'))
        style.configure('Small.TButton', font=(self.ui_font, 9))

    def build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=18, style='Shell.TFrame')
        main.pack(fill='both', expand=True)
        header = ttk.Frame(main, style='Shell.TFrame')
        header.pack(fill='x', pady=(0, 14))
        ttk.Label(header, text='\u842c\u7528\u8f49\u6a94\u738b', style='HeroTitle.TLabel').pack(anchor='w')
        ttk.Label(header, text='\u5716\u7247\u3001\u6587\u4ef6\u3001\u5f71\u97f3\u90fd\u80fd\u8f49\uff0c\u652f\u63f4 PDF\u3001Word\u3001Markdown\u3001\u5f71\u7247\u8207\u97f3\u8a0a\u683c\u5f0f\u3002', style='HeroSub.TLabel').pack(anchor='w', pady=(4, 2))

        stats = ttk.Frame(main, style='Card.TFrame', padding=14)
        stats.pack(fill='x', pady=(0, 14))
        self.tool_cards = {}
        for index, (label, key) in enumerate([('ffmpeg', 'ffmpeg'), ('pandoc', 'pandoc'), ('pdf2docx', 'pdf2docx')]):
            card = ttk.Frame(stats, style='Card.TFrame', padding=(10, 8))
            card.grid(row=0, column=index, sticky='nsew', padx=(0 if index == 0 else 10, 0))
            ttk.Label(card, text=label, style='Info.TLabel').pack(anchor='w')
            status_label = ttk.Label(card, text='', style='Metric.TLabel')
            status_label.pack(anchor='w', pady=(4, 2))
            source_label = ttk.Label(card, text='', style='Muted.TLabel')
            source_label.pack(anchor='w', pady=(0, 6))
            action_button = ttk.Button(card, text='', style='Small.TButton', command=lambda tool=key: self.handle_tool_button(tool))
            action_button.pack(anchor='w')
            self.tool_cards[key] = {'status': status_label, 'source': source_label, 'button': action_button}
            stats.columnconfigure(index, weight=1)
        ttk.Button(header, text='\u91cd\u65b0\u6574\u7406\u5de5\u5177\u72c0\u614b', command=self.refresh_tool_status, style='Small.TButton').pack(anchor='e')

        content = ttk.Frame(main, style='Shell.TFrame')
        content.pack(fill='both', expand=True)
        content.columnconfigure(0, weight=3)
        content.columnconfigure(1, weight=2)
        content.rowconfigure(0, weight=1)
        left = ttk.Frame(content, style='Shell.TFrame')
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 12))
        right = ttk.Frame(content, style='Shell.TFrame')
        right.grid(row=0, column=1, sticky='nsew')

        toolbar = ttk.LabelFrame(left, text='\u6a94\u6848\u64cd\u4f5c', padding=12, style='Section.TLabelframe')
        toolbar.pack(fill='x')
        ttk.Button(toolbar, text='\u9078\u64c7\u6a94\u6848', command=self.pick_files, style='Accent.TButton').pack(side='left')
        ttk.Button(toolbar, text='\u6e05\u7a7a\u5217\u8868', command=self.clear_files).pack(side='left', padx=8)
        ttk.Button(toolbar, text='\u9078\u64c7\u8f38\u51fa\u8cc7\u6599\u593e', command=self.pick_output_dir).pack(side='left')
        ttk.Label(toolbar, textvariable=self.drop_hint_var, style='Muted.TLabel').pack(side='right')

        options = ttk.LabelFrame(left, text='\u8f49\u6a94\u8a2d\u5b9a', padding=12, style='Section.TLabelframe')
        options.pack(fill='x', pady=(12, 12))
        ttk.Label(options, text='\u8f38\u51fa\u683c\u5f0f', style='Info.TLabel').grid(row=0, column=0, sticky='w')
        self.format_combo = ttk.Combobox(options, textvariable=self.format_var, state='readonly', width=18)
        self.format_combo.grid(row=0, column=1, sticky='w', padx=(10, 0))
        ttk.Label(options, text='\u8f38\u51fa\u8cc7\u6599\u593e', style='Info.TLabel').grid(row=1, column=0, sticky='w', pady=(10, 0))
        ttk.Entry(options, textvariable=self.output_dir_var, width=70).grid(row=1, column=1, sticky='we', padx=(10, 0), pady=(10, 0))
        options.columnconfigure(1, weight=1)

        file_frame = ttk.LabelFrame(left, text='\u5df2\u9078\u6a94\u6848', padding=10, style='Section.TLabelframe')
        file_frame.pack(fill='both', expand=True)
        self.file_list = tk.Listbox(file_frame, height=10, font=(self.mono_font, 10), bg='#fbfcff', fg='#16283f', selectbackground='#d8e7ff', selectforeground='#10233f', relief='flat', highlightthickness=0)
        self.file_list.pack(fill='both', expand=True)

        support_frame = ttk.LabelFrame(left, text='\u683c\u5f0f\u652f\u63f4', padding=10, style='Section.TLabelframe')
        support_frame.pack(fill='both', expand=True, pady=(12, 0))
        self.support_text = scrolledtext.ScrolledText(support_frame, height=9, wrap='word', bg='#f8fbff', fg='#10233f', relief='flat', highlightthickness=0, font=(self.mono_font, 10))
        self.support_text.pack(fill='both', expand=True)
        self.support_text.insert('1.0', '\u9078\u64c7\u6216\u62d6\u66f3\u6a94\u6848\u5f8c\uff0c\u9019\u88e1\u6703\u986f\u793a\u6bcf\u500b\u6a94\u6848\u76ee\u524d\u53ef\u8f49\u7684\u683c\u5f0f\u3002')
        self.support_text.configure(state='disabled')

        tips_frame = ttk.LabelFrame(right, text='\u4f7f\u7528\u63d0\u793a', padding=12, style='Section.TLabelframe')
        tips_frame.pack(fill='x')
        for tip in ['\u53ef\u76f4\u63a5\u628a\u6a94\u6848\u62d6\u9032\u5de6\u5074\u6e05\u55ae\u3002', '\u683c\u5f0f\u652f\u63f4\u5340\u6703\u5217\u51fa\u6bcf\u500b\u6a94\u6848\u76ee\u524d\u53ef\u8f49\u7684\u683c\u5f0f\u3002', 'PDF \u73fe\u5728\u53ef\u8f49\u6210 docx\u3001txt\u3001md\u3002', '\u9032\u5ea6\u689d\u4e0b\u65b9\u6703\u76f4\u63a5\u986f\u793a\u57f7\u884c\u7d00\u9304\u8207\u547d\u4ee4\u8f38\u51fa\u3002']:
            ttk.Label(tips_frame, text='- ' + tip, style='Info.TLabel').pack(anchor='w', pady=2)

        action_frame = ttk.LabelFrame(right, text='\u57f7\u884c', padding=12, style='Section.TLabelframe')
        action_frame.pack(fill='both', expand=True, pady=(12, 0))
        ttk.Button(action_frame, text='\u958b\u59cb\u6279\u6b21\u8f49\u6a94', command=self.start_conversion, style='Accent.TButton').pack(anchor='w')
        ttk.Label(action_frame, textvariable=self.status_var, style='Info.TLabel').pack(anchor='w', pady=(10, 0))
        ttk.Progressbar(action_frame, variable=self.progress_value, maximum=100, mode='determinate').pack(fill='x', pady=(10, 4))
        ttk.Label(action_frame, textvariable=self.progress_text_var, style='Muted.TLabel').pack(anchor='w')

        ttk.Label(action_frame, text='\u57f7\u884c\u7d00\u9304', style='Info.TLabel').pack(anchor='w', pady=(12, 6))
        self.log_text = tk.Text(action_frame, height=14, wrap='word', bg='#0f1d2e', fg='#eef5ff', insertbackground='#eef5ff', relief='flat', highlightthickness=0, font=(self.mono_font, 10))
        self.log_text.pack(fill='both', expand=True)
        self.refresh_tool_status()

    def is_tool_installed(self, tool_name: str) -> bool:
        return {
            'ffmpeg': self.converter.tools.ffmpeg,
            'pandoc': self.converter.tools.pandoc,
            'pdf2docx': PdfToDocxConverter is not None,
        }.get(tool_name, False)

    def get_tool_meta(self, tool_name: str) -> tuple[str, str]:
        if tool_name == 'ffmpeg':
            source = self.converter.tools.ffmpeg_source
        elif tool_name == 'pandoc':
            source = self.converter.tools.pandoc_source
        elif tool_name == 'pdf2docx':
            return ('\u5df2\u5167\u5efa Python \u5957\u4ef6' if PdfToDocxConverter else '\u672a\u5b89\u88dd', 'bundled' if PdfToDocxConverter else 'missing')
        else:
            return ('\u672a\u77e5', 'missing')
        return ({'embedded': '\u5167\u5efa\u74b0\u5883', 'portable': 'tools \u8cc7\u6599\u593e', 'missing': '\u672a\u5b89\u88dd'}.get(source, source), source)

    def refresh_tool_status(self) -> None:
        self.converter = SuperConverter()
        for tool_name, widgets in self.tool_cards.items():
            installed = self.is_tool_installed(tool_name)
            source_text, source_key = self.get_tool_meta(tool_name)
            widgets['status'].configure(text='\u5df2\u5b89\u88dd' if installed else '\u672a\u5b89\u88dd')
            widgets['source'].configure(text=source_text)
            if tool_name == 'pdf2docx':
                widgets['button'].configure(text='\u5167\u5efa\u5957\u4ef6', state='disabled')
            elif installed and source_key in {'portable', 'embedded', 'bundled'}:
                widgets['button'].configure(text='\u958b\u555f tools', state='normal')
            else:
                widgets['button'].configure(text='\u67e5\u770b\u4f4d\u7f6e', state='normal')
        self.append_log('\u5df2\u91cd\u65b0\u6574\u7406\u5de5\u5177\u72c0\u614b')

    def handle_tool_button(self, tool_name: str) -> None:
        tools_dir = app_root() / 'tools'
        target = {
            'ffmpeg': tools_dir / 'ffmpeg',
            'pandoc': tools_dir / 'pandoc',
        }.get(tool_name, tools_dir)
        target.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(['explorer', str(target)])

    def pick_files(self) -> None:
        paths = filedialog.askopenfilenames(title='\u9078\u64c7\u8981\u8f49\u63db\u7684\u6a94\u6848', filetypes=[('All files', '*.*')])
        if not paths:
            return
        self.add_files([Path(path) for path in paths], replace=True)

    def add_files(self, paths: list[Path], replace: bool = False) -> None:
        incoming = [Path(path) for path in paths if Path(path).exists() and Path(path).is_file()]
        merged = [] if replace else list(self.selected_files)
        seen = {str(path).lower() for path in merged}
        added = 0
        for path in incoming:
            key = str(path).lower()
            if key in seen:
                continue
            merged.append(path)
            seen.add(key)
            added += 1
        self.selected_files = merged
        self.refresh_file_list()
        self.refresh_output_formats()
        if added:
            self.append_log(f'\u5df2\u52a0\u5165 {added} \u500b\u6a94\u6848')

    def clear_files(self) -> None:
        self.selected_files = []
        self.file_output_map = {}
        self.file_list.delete(0, tk.END)
        self.format_combo['values'] = []
        self.format_var.set('')
        self.progress_value.set(0)
        self.progress_text_var.set('0 / 0')
        self.status_var.set('\u5df2\u6e05\u7a7a\u6a94\u6848\u5217\u8868')
        self.update_support_text([], {})

    def pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title='\u9078\u64c7\u8f38\u51fa\u8cc7\u6599\u593e')
        if path:
            self.output_dir_var.set(path)

    def refresh_file_list(self) -> None:
        self.file_list.delete(0, tk.END)
        for path in self.selected_files:
            self.file_list.insert(tk.END, str(path))
        self.status_var.set(f'\u5df2\u9078\u64c7 {len(self.selected_files)} \u500b\u6a94\u6848')
        self.progress_text_var.set(f'0 / {len(self.selected_files)}')

    def refresh_output_formats(self) -> None:
        if not self.selected_files:
            self.update_support_text([], {})
            return
        common = None
        per_file: dict[Path, list[str]] = {}
        for path in self.selected_files:
            outputs = sorted(self.converter.get_supported_outputs(path))
            per_file[path] = outputs
            output_set = set(outputs)
            common = output_set if common is None else common & output_set
        self.file_output_map = per_file
        formats = sorted(common or [])
        self.format_combo['values'] = formats
        if formats:
            self.format_var.set(formats[0])
            self.status_var.set(f'\u5df2\u9078\u64c7 {len(self.selected_files)} \u500b\u6a94\u6848\uff0c\u53ef\u5171\u540c\u8f38\u51fa {len(formats)} \u7a2e\u683c\u5f0f')
        else:
            self.format_var.set('')
            self.status_var.set('\u9019\u6279\u6a94\u6848\u6c92\u6709\u5171\u540c\u53ef\u7528\u7684\u8f38\u51fa\u683c\u5f0f')
            self.append_log('\u9019\u6279\u6a94\u6848\u6c92\u6709\u5171\u540c\u53ef\u7528\u7684\u8f38\u51fa\u683c\u5f0f\uff0c\u8acb\u5206\u958b\u8655\u7406\u4e0d\u540c\u985e\u578b\u6a94\u6848\u3002')
        self.update_support_text(formats, per_file)

    def update_support_text(self, common_formats: list[str], per_file: dict[Path, list[str]]) -> None:
        self.support_text.configure(state='normal')
        self.support_text.delete('1.0', tk.END)
        if not per_file:
            self.support_text.insert(tk.END, '\u9078\u64c7\u6216\u62d6\u66f3\u6a94\u6848\u5f8c\uff0c\u9019\u88e1\u6703\u986f\u793a\u6bcf\u500b\u6a94\u6848\u76ee\u524d\u53ef\u8f49\u7684\u683c\u5f0f\u3002')
            self.support_text.configure(state='disabled')
            return
        if common_formats:
            self.support_text.insert(tk.END, '\u5171\u540c\u53ef\u8f38\u51fa\u683c\u5f0f\\n')
            self.support_text.insert(tk.END, '  ' + ', '.join(common_formats) + '\\n\\n')
        else:
            self.support_text.insert(tk.END, '\u5171\u540c\u53ef\u8f38\u51fa\u683c\u5f0f\\n')
            self.support_text.insert(tk.END, '  \u6c92\u6709\u5171\u540c\u683c\u5f0f\uff0c\u5efa\u8b70\u5206\u958b\u6279\u6b21\u8655\u7406\u3002\\n\\n')
        self.support_text.insert(tk.END, '\u5404\u6a94\u6848\u652f\u63f4\u683c\u5f0f\\n')
        for path, outputs in per_file.items():
            if outputs:
                self.support_text.insert(tk.END, f'- {path.name}\\n  ' + ', '.join(outputs) + '\\n')
            else:
                self.support_text.insert(tk.END, f'- {path.name}\\n  \u76ee\u524d\u6c92\u6709\u53ef\u7528\u8f38\u51fa\u683c\u5f0f\\n')
        self.support_text.configure(state='disabled')
    def start_conversion(self) -> None:
        if not self.selected_files:
            messagebox.showwarning('\u63d0\u9192', '\u8acb\u5148\u9078\u64c7\u6a94\u6848')
            return
        target_ext = self.format_var.get().strip()
        if not target_ext:
            messagebox.showwarning('\u63d0\u9192', '\u8acb\u5148\u9078\u64c7\u8f38\u51fa\u683c\u5f0f')
            return
        output_dir = Path(self.output_dir_var.get())
        output_dir.mkdir(parents=True, exist_ok=True)
        self.progress_value.set(0)
        self.progress_text_var.set(f'0 / {len(self.selected_files)}')
        self.status_var.set('\u8f49\u6a94\u9032\u884c\u4e2d...')
        threading.Thread(target=self.run_batch_conversion, args=(output_dir, target_ext), daemon=True).start()

    def run_batch_conversion(self, output_dir: Path, target_ext: str) -> None:
        total = len(self.selected_files)
        success = 0
        self.append_log(f'\u6279\u6b21\u8f49\u6a94\u958b\u59cb\uff0c\u5171 {total} \u500b\u6a94\u6848\uff0c\u8f38\u51fa\u683c\u5f0f {target_ext}')
        for index, source in enumerate(self.selected_files, start=1):
            target = output_dir / f'{source.stem}{target_ext}'
            try:
                self.append_log(f'[{index}/{total}] \u6e96\u5099\u8f49\u63db {source.name}')
                self.converter.convert(source, target, self.append_log)
                self.append_log(f'[{index}/{total}] \u5b8c\u6210\uff0c\u8f38\u51fa\u6a94\u6848: {target}')
                success += 1
            except Exception as exc:
                self.append_log(f'[\u5931\u6557] {source.name}: {exc}')
            progress = (index / total) * 100 if total else 0
            self.root.after(0, lambda i=index, t=total, p=progress: self.update_progress(i, t, p))
        self.root.after(0, lambda: self.status_var.set(f'\u5b8c\u6210 {success}/{total} \u500b\u6a94\u6848'))
        self.append_log('\u5168\u90e8\u5de5\u4f5c\u5df2\u5b8c\u6210')

    def update_progress(self, current: int, total: int, percent: float) -> None:
        self.progress_value.set(percent)
        self.progress_text_var.set(f'{current} / {total} ({percent:.0f}%)')

    def enable_file_drop(self) -> None:
        self.drop_hint_var.set('\u62d6\u66f3\u52a0\u5165\u6a94\u6848\u529f\u80fd\u66ab\u6642\u95dc\u9589\uff0c\u8acb\u4f7f\u7528\u300c\u9078\u64c7\u6a94\u6848\u300d')

    def _extract_drop_files(self, hdrop: int) -> list[Path]:
        count = self._shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
        paths: list[Path] = []
        for index in range(count):
            length = self._shell32.DragQueryFileW(hdrop, index, None, 0)
            buffer = ctypes.create_unicode_buffer(length + 1)
            self._shell32.DragQueryFileW(hdrop, index, buffer, length + 1)
            paths.append(Path(buffer.value))
        self._shell32.DragFinish(hdrop)
        return paths

    def append_log(self, message: str) -> None:
        def write() -> None:
            timestamp = datetime.now().strftime('%H:%M:%S')
            self.log_text.insert(tk.END, f'[{timestamp}] {message}\n')
            self.log_text.see(tk.END)
        self.root.after(0, write)

def main() -> None:
    write_startup_trace('main:enter')
    root = tk.Tk()
    write_startup_trace('main:tk_created')
    app = App(root)
    write_startup_trace('main:app_created')
    if len(sys.argv) > 1:
        app.selected_files = [Path(arg) for arg in sys.argv[1:] if Path(arg).exists()]
        app.refresh_file_list()
        app.refresh_output_formats()
    write_startup_trace('main:before_mainloop')
    root.mainloop()
    write_startup_trace('main:after_mainloop')


if __name__ == '__main__':
    main()







