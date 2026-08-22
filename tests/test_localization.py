import unittest
import tempfile
from pathlib import Path
from unittest import mock

import pymupdf

import ScanItEasy_backend as backend
import ScanItEasy_GUI as gui


class LocalizationTests(unittest.TestCase):
    def test_application_version(self):
        self.assertEqual(gui.APP_VERSION, "2.1.0")

    def test_russian_and_english_have_the_same_keys(self):
        self.assertEqual(set(gui.STRINGS["ru"]), set(gui.STRINGS["en"]))

    def test_default_mode_is_not_exposed(self):
        self.assertNotIn("default_mode", dict(gui.MODES))
        self.assertEqual(gui.MODES[0][0], "custom_mode")

    def test_docx_image_modes_are_exposed_next_to_docx_pdf(self):
        modes = [mode for mode, _label in gui.MODES]
        self.assertEqual(
            modes[1:4],
            ["convert_to_pdf_mode", "convert_docx_to_png_mode", "convert_docx_to_jpg_mode"],
        )

    def test_backend_returns_language_neutral_error_code(self):
        result = backend.process_job({"mode": "not-a-mode"})
        self.assertFalse(result.success)
        self.assertEqual(result.message, "unknown_mode")

    def test_guide_links_exist_for_each_language(self):
        self.assertEqual(set(gui.GUIDE_URLS), set(gui.STRINGS))
        self.assertEqual(set(gui.GHOSTSCRIPT_GUIDE_URLS), set(gui.STRINGS))

    def test_only_multipage_pdf_enables_last_part_page_selection(self):
        with tempfile.TemporaryDirectory() as folder:
            one_page = Path(folder) / "one.pdf"
            two_pages = Path(folder) / "two.pdf"
            for path, count in ((one_page, 1), (two_pages, 2)):
                with pymupdf.open() as document:
                    for _ in range(count):
                        document.new_page()
                    document.save(path)
            self.assertFalse(gui.is_multipage_pdf(str(one_page)))
            self.assertTrue(gui.is_multipage_pdf(str(two_pages)))
            self.assertFalse(gui.is_multipage_pdf("last-page.png"))


class ImageSelectionTests(unittest.TestCase):
    def test_new_image_selection_is_appended_without_duplicates(self):
        class FakeApp:
            selected_images = ["C:/first/one.png"]
            rendered = False

            @staticmethod
            def file_types(_kind):
                return []

            def render_image_rows(self):
                self.rendered = True

        app = FakeApp()
        with mock.patch.object(
            gui.filedialog,
            "askopenfilenames",
            return_value=("C:/second/two.jpg", "C:/first/one.png"),
        ):
            gui.ScanItEasyApp.select_images(app)

        self.assertEqual(
            app.selected_images,
            ["C:/first/one.png", "C:/second/two.jpg"],
        )
        self.assertTrue(app.rendered)


if __name__ == "__main__":
    unittest.main()
