import os
import tempfile
import unittest
from unittest import mock

from core.converter import (
    BACKEND_LIBREOFFICE,
    BACKEND_PPT_COM,
    BACKEND_PPT_MAC,
    BACKEND_WORD_COM,
    BACKEND_WORD_MAC,
    _applescript_string,
    backends_for_file,
    convert_one_with_fallback,
)


class ConverterBackendSelectionTest(unittest.TestCase):
    def test_ppt_file_never_routes_to_word_backend(self):
        backends = [
            BACKEND_WORD_MAC,
            BACKEND_PPT_MAC,
            BACKEND_WORD_COM,
            BACKEND_PPT_COM,
            BACKEND_LIBREOFFICE,
        ]
        self.assertEqual(
            backends_for_file("demo.pptx", backends),
            [BACKEND_PPT_MAC, BACKEND_PPT_COM, BACKEND_LIBREOFFICE],
        )

    def test_word_file_never_routes_to_powerpoint_backend(self):
        backends = [
            BACKEND_PPT_MAC,
            BACKEND_WORD_MAC,
            BACKEND_PPT_COM,
            BACKEND_WORD_COM,
            BACKEND_LIBREOFFICE,
        ]
        self.assertEqual(
            backends_for_file("notes.docx", backends),
            [BACKEND_WORD_MAC, BACKEND_WORD_COM, BACKEND_LIBREOFFICE],
        )

    def test_applescript_string_escapes_quotes_and_backslashes(self):
        self.assertEqual(
            _applescript_string('/tmp/a "quoted" \\ file.docx'),
            '"/tmp/a \\"quoted\\" \\\\ file.docx"',
        )


class ConverterFallbackBehaviorTest(unittest.TestCase):
    def test_prefetched_pdf_render_failure_falls_back_to_libreoffice(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_dir = os.path.join(tmpdir, "demo")
            logs = []
            with mock.patch(
                "core.converter._pdf_to_png",
                side_effect=RuntimeError("render failed"),
            ) as pdf_mock, mock.patch(
                "core.converter._convert_libreoffice",
                return_value=6,
            ) as libre_mock:
                pages, used, actual_out_dir = convert_one_with_fallback(
                    "demo.pptx",
                    target_dir,
                    17,
                    [BACKEND_PPT_MAC, BACKEND_LIBREOFFICE],
                    soffice_path="/fake/soffice",
                    pdf_path="/tmp/prefetched.pdf",
                    pdf_backend=BACKEND_PPT_MAC,
                    log=logs.append,
                )

            self.assertEqual(pages, 6)
            self.assertEqual(used, BACKEND_LIBREOFFICE)
            self.assertEqual(actual_out_dir, target_dir)
            self.assertTrue(os.path.isdir(actual_out_dir))
            pdf_mock.assert_called_once()
            libre_mock.assert_called_once()
            self.assertTrue(
                any("预导 PDF 渲染失败" in entry for entry in logs),
                logs,
            )

    def test_failure_cleans_temp_dir_and_does_not_leave_final_dir(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_dir = os.path.join(tmpdir, "broken")
            with mock.patch(
                "core.converter._convert_libreoffice",
                side_effect=RuntimeError("libre failed"),
            ):
                with self.assertRaisesRegex(RuntimeError, "所有转换引擎均失败"):
                    convert_one_with_fallback(
                        "broken.docx",
                        target_dir,
                        17,
                        [BACKEND_LIBREOFFICE],
                        soffice_path="/fake/soffice",
                    )

            self.assertFalse(os.path.exists(target_dir))
            leftovers = [
                name for name in os.listdir(tmpdir)
                if name.startswith(".broken.tmp_")
            ]
            self.assertEqual(leftovers, [])


if __name__ == "__main__":
    unittest.main()
