import ctypes
import sys
import traceback
from pathlib import Path


def write_startup_error_log(error_text: str) -> Path:
    desktop = Path.home() / "Desktop"
    if not desktop.exists():
        desktop = Path.home()
    log_path = desktop / "wanyong_converter_startup_error.txt"
    log_path.write_text(error_text, encoding="utf-8")
    return log_path


def write_startup_trace(message: str) -> None:
    desktop = Path.home() / "Desktop"
    if not desktop.exists():
        desktop = Path.home()
    trace_path = desktop / "wanyong_converter_startup_trace.txt"
    with trace_path.open("a", encoding="utf-8") as handle:
        handle.write(message + "\n")


def show_error_dialog(message: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, message, "Wanyong Converter Startup Error", 0x10)
    except Exception:
        pass


def run() -> None:
    write_startup_trace("launcher:start")
    from super_converter import main
    write_startup_trace("launcher:imported_main")

    write_startup_trace("launcher:before_main")
    main()
    write_startup_trace("launcher:after_main")


if __name__ == "__main__":
    try:
        write_startup_trace("launcher:__main__")
        run()
    except Exception:
        details = "".join(traceback.format_exception(*sys.exc_info()))
        log_path = write_startup_error_log(details)
        write_startup_trace("launcher:exception")
        show_error_dialog(f"Startup failed.\\n\\nLog file:\\n{log_path}")
        raise
