import os
import tempfile
import unittest

from core.scanner import scan_ppt_files, scan_supported_files


class ScannerTest(unittest.TestCase):
    def test_scan_supported_files_includes_ppt_and_word_and_skips_temp(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            os.makedirs(os.path.join(tmpdir, "nested"), exist_ok=True)
            paths = [
                os.path.join(tmpdir, "demo.pptx"),
                os.path.join(tmpdir, "nested", "plan.docx"),
                os.path.join(tmpdir, "nested", "~$draft.docx"),
                os.path.join(tmpdir, "ignore.txt"),
            ]
            for path in paths:
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write("x")

            results = scan_supported_files(tmpdir)

            self.assertEqual(
                results,
                sorted(
                    [
                        os.path.join(tmpdir, "demo.pptx"),
                        os.path.join(tmpdir, "nested", "plan.docx"),
                    ]
                ),
            )

    def test_scan_ppt_files_keeps_legacy_behavior(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            ppt_path = os.path.join(tmpdir, "slides.ppt")
            doc_path = os.path.join(tmpdir, "notes.doc")
            for path in [ppt_path, doc_path]:
                with open(path, "w", encoding="utf-8") as handle:
                    handle.write("x")

            self.assertEqual(scan_ppt_files(tmpdir), [ppt_path])


if __name__ == "__main__":
    unittest.main()
