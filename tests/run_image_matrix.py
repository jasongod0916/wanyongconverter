import tempfile
from collections import defaultdict
from pathlib import Path
import sys

from PIL import Image

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from super_converter import SuperConverter


def noop_log(_: str) -> None:
    return


def make_source_image(ext: str, target: Path) -> None:
    if ext == ".svg":
        target.write_text(
            """<svg xmlns="http://www.w3.org/2000/svg" width="96" height="64" viewBox="0 0 96 64">
<rect width="96" height="64" fill="#1f5fbf" />
<circle cx="22" cy="22" r="12" fill="#ffd166" />
<rect x="42" y="14" width="40" height="12" rx="3" fill="#06d6a0" />
<text x="42" y="40" font-size="12" font-family="Arial, sans-serif" fill="#ffffff">SVG</text>
</svg>
""",
            encoding="utf-8",
        )
        return

    if ext == ".jpg":
        img = Image.new("RGB", (96, 64), color=(40, 120, 210))
        img.save(target, quality=95)
        return

    if ext == ".bmp":
        img = Image.new("RGB", (96, 64), color=(220, 180, 60))
        img.save(target)
        return

    if ext == ".gif":
        img = Image.new("P", (96, 64))
        palette = []
        for value in range(256):
            palette.extend((value, 255 - value, min(255, value * 2)))
        img.putpalette(palette[:768])
        for x in range(96):
            for y in range(64):
                img.putpixel((x, y), (x + y) % 255)
        img.save(target)
        return

    if ext == ".ico":
        img = Image.new("RGBA", (64, 64), color=(0, 0, 0, 0))
        for x in range(64):
            for y in range(64):
                img.putpixel((x, y), (x * 4 % 255, y * 4 % 255, 180, 255))
        img.save(target, sizes=[(64, 64), (32, 32), (16, 16)])
        return

    if ext == ".webp":
        img = Image.new("RGBA", (96, 64), color=(90, 200, 120, 180))
        img.save(target, format="WEBP")
        return

    if ext in {".tif", ".tiff"}:
        img = Image.new("RGB", (96, 64), color=(140, 80, 190))
        img.save(target, format="TIFF")
        return

    img = Image.new("RGBA", (96, 64), color=(230, 70, 90, 180))
    img.save(target)


def verify_output(target_path: Path) -> None:
    if target_path.suffix.lower() == ".pdf":
        with target_path.open("rb") as handle:
            header = handle.read(4)
        if header != b"%PDF":
            raise ValueError("output is not a valid PDF header")
        return

    with Image.open(target_path) as img:
        img.load()
        if img.width <= 0 or img.height <= 0:
            raise ValueError("output image has invalid size")


def main() -> None:
    converter = SuperConverter()
    source_exts = [".png", ".jpg", ".bmp", ".gif", ".tiff", ".webp", ".ico", ".svg"]

    results: list[dict] = []
    failures_by_target: dict[str, list[str]] = defaultdict(list)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)

        source_paths: list[Path] = []
        for ext in source_exts:
            source_path = temp_root / f"sample{ext}"
            make_source_image(ext, source_path)
            source_paths.append(source_path)

        for source_path in source_paths:
            supported_outputs = converter.get_supported_outputs(source_path)
            for target_ext in supported_outputs:
                target_path = temp_root / f"{source_path.stem}_to_{target_ext.lstrip('.')}{target_ext}"
                result = {
                    "source_ext": source_path.suffix.lower(),
                    "target_ext": target_ext,
                    "status": "success",
                    "error": "",
                }
                try:
                    converter.convert(source_path, target_path, noop_log)
                    if not target_path.exists() or target_path.stat().st_size == 0:
                        raise ValueError("output file missing or empty")
                    verify_output(target_path)
                except Exception as exc:
                    result["status"] = "failed"
                    result["error"] = str(exc)
                    failures_by_target[target_ext].append(source_path.suffix.lower())
                results.append(result)

    summary = {
        "attempted": len(results),
        "success": sum(1 for item in results if item["status"] == "success"),
        "failed": sum(1 for item in results if item["status"] == "failed"),
        "failed_targets": dict(failures_by_target),
    }

    print("=== Image Conversion Matrix Summary ===")
    print(f"Attempted: {summary['attempted']}")
    print(f"Success:   {summary['success']}")
    print(f"Failed:    {summary['failed']}")
    print("")
    print("=== Successful Pairs ===")
    for item in results:
        if item["status"] == "success":
            print(f"{item['source_ext']} -> {item['target_ext']}")
    print("")
    print("=== Failed Pairs ===")
    for item in results:
        if item["status"] == "failed":
            print(f"{item['source_ext']} -> {item['target_ext']} | {item['error']}")


if __name__ == "__main__":
    main()
