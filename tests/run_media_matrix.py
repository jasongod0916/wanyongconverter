import subprocess
import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from super_converter import SuperConverter  # noqa: E402


VIDEO_SOURCE_EXTS = [
    ".mp4",
    ".mov",
    ".avi",
    ".mkv",
    ".webm",
    ".flv",
    ".wmv",
]

AUDIO_SOURCE_EXTS = [
    ".mp3",
    ".wav",
    ".aac",
    ".flac",
    ".ogg",
    ".m4a",
]


def noop_log(_: str) -> None:
    return


def ffmpeg_cmd(converter: SuperConverter) -> str:
    return converter.tools.ffmpeg_cmd or "ffmpeg"


def run_ffmpeg(converter: SuperConverter, args: list[str]) -> None:
    cmd = [ffmpeg_cmd(converter), "-y", *args]
    subprocess.run(cmd, capture_output=True, check=True)


def create_base_sources(converter: SuperConverter, temp_root: Path) -> tuple[Path, Path]:
    base_video = temp_root / "base_video.mp4"
    base_audio = temp_root / "base_audio.wav"

    run_ffmpeg(
        converter,
        [
            "-f",
            "lavfi",
            "-i",
            "testsrc=size=320x180:rate=24:duration=2",
            "-f",
            "lavfi",
            "-i",
            "sine=frequency=880:duration=2",
            "-shortest",
            "-c:v",
            "libx264",
            "-pix_fmt",
            "yuv420p",
            "-c:a",
            "aac",
            str(base_video),
        ],
    )

    run_ffmpeg(
        converter,
        [
            "-f",
            "lavfi",
            "-i",
            "sine=frequency=660:duration=2",
            "-c:a",
            "pcm_s16le",
            str(base_audio),
        ],
    )

    return base_video, base_audio


def build_video_args(target: Path) -> list[str]:
    ext = target.suffix.lower()
    if ext == ".mp4":
        return ["-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "aac", str(target)]
    if ext == ".mov":
        return ["-c:v", "mpeg4", "-c:a", "aac", str(target)]
    if ext == ".avi":
        return ["-c:v", "mpeg4", "-c:a", "mp3", str(target)]
    if ext == ".mkv":
        return ["-c:v", "libx264", "-c:a", "aac", str(target)]
    if ext == ".webm":
        return ["-c:v", "libvpx-vp9", "-c:a", "libopus", str(target)]
    if ext == ".flv":
        return ["-c:v", "flv", "-c:a", "libmp3lame", str(target)]
    if ext == ".wmv":
        return ["-c:v", "wmv2", "-c:a", "wmav2", str(target)]
    return [str(target)]


def build_audio_args(target: Path) -> list[str]:
    ext = target.suffix.lower()
    if ext == ".mp3":
        return ["-vn", "-c:a", "libmp3lame", "-q:a", "2", str(target)]
    if ext == ".wav":
        return ["-vn", "-c:a", "pcm_s16le", str(target)]
    if ext == ".aac":
        return ["-vn", "-c:a", "aac", str(target)]
    if ext == ".flac":
        return ["-vn", "-c:a", "flac", str(target)]
    if ext == ".ogg":
        return ["-vn", "-c:a", "libvorbis", str(target)]
    if ext == ".m4a":
        return ["-vn", "-c:a", "aac", str(target)]
    return [str(target)]


def create_source_files(converter: SuperConverter, temp_root: Path) -> dict[str, Path]:
    base_video, base_audio = create_base_sources(converter, temp_root)
    sources: dict[str, Path] = {}

    for ext in VIDEO_SOURCE_EXTS:
        target = temp_root / f"sample{ext}"
        run_ffmpeg(converter, ["-i", str(base_video), *build_video_args(target)])
        sources[ext] = target

    for ext in AUDIO_SOURCE_EXTS:
        target = temp_root / f"sample{ext}"
        run_ffmpeg(converter, ["-i", str(base_audio), *build_audio_args(target)])
        sources[ext] = target

    return sources


def verify_media_output(converter: SuperConverter, target: Path) -> None:
    if not target.exists() or target.stat().st_size == 0:
        raise ValueError("output file missing or empty")

    cmd = [
        ffmpeg_cmd(converter),
        "-v",
        "error",
        "-i",
        str(target),
        "-f",
        "null",
        "-",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise ValueError(result.stderr.strip() or "ffmpeg could not decode the output")


def main() -> None:
    converter = SuperConverter()
    if not converter.tools.ffmpeg:
        raise RuntimeError("ffmpeg is not available")

    results: list[dict[str, str]] = []

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        sources = create_source_files(converter, temp_root)

        print("=== Media Source Files ===")
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
                    verify_media_output(converter, target_path)
                except Exception as exc:
                    result["status"] = "failed"
                    result["error"] = str(exc)
                results.append(result)

    success_items = [item for item in results if item["status"] == "success"]
    failed_items = [item for item in results if item["status"] == "failed"]

    print("=== Media Conversion Matrix Summary ===")
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
