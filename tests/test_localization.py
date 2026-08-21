import unittest

import ScanItEasy_backend as backend
import ScanItEasy_GUI as gui


class LocalizationTests(unittest.TestCase):
    def test_application_version(self):
        self.assertEqual(gui.APP_VERSION, "2.0.0")

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


if __name__ == "__main__":
    unittest.main()
