import tempfile
import unittest
from pathlib import Path

from PIL import Image

from super_converter import SuperConverter


def noop_log(_: str) -> None:
    return


class TestImageConversion(unittest.TestCase):
    # 這組來源格式盡量涵蓋一般使用者常見的圖片類型。
    SOURCE_CANDIDATES = [".png", ".jpg", ".bmp", ".gif", ".tiff", ".webp", ".ico"]
    TARGET_CANDIDATES = [".png", ".jpg", ".bmp", ".webp", ".tiff", ".pdf"]

    @classmethod
    def setUpClass(cls) -> None:
        cls.converter = SuperConverter()

    def test_common_image_inputs_are_listed(self) -> None:
        inputs = set(self.converter.get_image_inputs())
        for ext in [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".webp", ".ico"]:
            self.assertIn(ext, inputs)

    def test_common_image_outputs_are_listed(self) -> None:
        outputs = set(self.converter.get_image_outputs())
        for ext in [".png", ".jpg", ".bmp", ".pdf"]:
            self.assertIn(ext, outputs)

    def test_generated_images_can_convert_to_common_targets(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            generated_sources = self._create_source_images(temp_root)

            self.assertTrue(generated_sources, "沒有成功建立任何測試圖片")

            for source_path in generated_sources:
                supported_outputs = set(self.converter.get_supported_outputs(source_path))
                planned_targets = [
                    ext
                    for ext in self.TARGET_CANDIDATES
                    if ext != source_path.suffix.lower() and ext in supported_outputs
                ]

                self.assertTrue(
                    planned_targets,
                    f"{source_path.suffix} 沒有找到可測的目標格式",
                )

                for target_ext in planned_targets:
                    with self.subTest(source=source_path.suffix.lower(), target=target_ext):
                        target_path = temp_root / f"{source_path.stem}_to_{target_ext.lstrip('.')}{target_ext}"
                        self.converter.convert(source_path, target_path, noop_log)
                        self.assertTrue(target_path.exists(), f"轉檔後未產生輸出：{target_path.name}")
                        self.assertGreater(target_path.stat().st_size, 0, f"輸出檔案為空：{target_path.name}")
                        self._assert_output_is_readable(target_path)

    def _create_source_images(self, temp_root: Path) -> list[Path]:
        created: list[Path] = []

        for ext in self.SOURCE_CANDIDATES:
            target = temp_root / f"sample{ext}"
            try:
                self._save_sample_image(ext, target)
            except Exception:
                # 某些 Pillow build 不一定支援所有格式，這裡直接略過不可建立的格式。
                continue
            created.append(target)

        return created

    def _save_sample_image(self, ext: str, target: Path) -> None:
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

        if ext == ".tiff":
            img = Image.new("RGB", (96, 64), color=(140, 80, 190))
            img.save(target, format="TIFF")
            return

        img = Image.new("RGBA", (96, 64), color=(230, 70, 90, 180))
        img.save(target)

    def _assert_output_is_readable(self, target_path: Path) -> None:
        if target_path.suffix.lower() == ".pdf":
            with target_path.open("rb") as handle:
                header = handle.read(4)
            self.assertEqual(header, b"%PDF")
            return

        with Image.open(target_path) as img:
            img.load()
            self.assertGreater(img.width, 0)
            self.assertGreater(img.height, 0)


if __name__ == "__main__":
    unittest.main()
