import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pymupdf
from PIL import Image

import ScanItEasy_backend as backend


class PageSequenceTests(unittest.TestCase):
    def test_order_reverse_ranges_and_duplicates_are_preserved(self):
        self.assertEqual(
            backend.parse_page_sequence("7, 1-5, 5-1, 3, 1-3"),
            [7, 1, 2, 3, 4, 5, 5, 4, 3, 2, 1, 3, 1, 2, 3],
        )

    def test_zero_and_negative_pages_are_rejected(self):
        for value in ("0", "-1", "1, 0", "0-3"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                backend.parse_page_sequence(value)

    def test_out_of_range_pages_are_skipped(self):
        selected, skipped = backend.select_pages("5, 8, 3-6", 5)
        self.assertEqual(selected, [4, 2, 3, 4])
        self.assertEqual(skipped, [8, 6])


class ConversionTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.folder = Path(self.temporary.name)
        self.png = self.folder / "one.png"
        self.jpg = self.folder / "two.jpg"
        Image.new("RGBA", (120, 80), (255, 0, 0, 128)).save(self.png)
        Image.new("RGB", (80, 120), "blue").save(self.jpg, quality=90)

    def tearDown(self):
        self.temporary.cleanup()

    def test_multiple_mixed_images_become_one_pdf(self):
        output = self.folder / "images.pdf"
        backend.images_to_pdf([str(self.png), str(self.jpg)], output)
        with pymupdf.open(output) as document:
            self.assertEqual(document.page_count, 2)

    def test_pdf_page_order_and_duplicates_are_preserved(self):
        source = self.folder / "source.pdf"
        backend.images_to_pdf([str(self.png), str(self.jpg)], source)
        output = self.folder / "selected.pdf"
        backend.copy_pages(source, output, "2, 1, 2")
        with pymupdf.open(output) as document:
            self.assertEqual(document.page_count, 3)

    def test_pdf_image_names_use_source_page_and_number_duplicates(self):
        source = self.folder / "Перевод паспорта.pdf"
        backend.images_to_pdf([str(self.png), str(self.jpg)], source)

        outputs, skipped = backend.pdf_to_images(source, "png", "1, 2, 1", 72, 85)

        self.assertEqual(skipped, [])
        self.assertEqual(
            [Path(output).name for output in outputs],
            [
                "Перевод паспорта_page_001.png",
                "Перевод паспорта_page_002.png",
                "Перевод паспорта_page_001(2).png",
            ],
        )
        for output in outputs:
            with Image.open(output) as image:
                image.verify()

    def test_docx_to_images_uses_docx_name_and_page_rules(self):
        source = self.folder / "document.docx"
        source.write_bytes(b"placeholder")

        def fake_convert(_source, output):
            backend.images_to_pdf([str(self.png), str(self.jpg)], output)

        with mock.patch.object(backend, "is_docx_open_in_word", return_value=False), \
             mock.patch.object(backend, "convert_docx", side_effect=fake_convert):
            result = backend.process_job({
                "mode": "convert_docx_to_jpg_mode",
                "docx_path": str(source),
                "pages": "2, 1, 2",
                "dpi": 72,
                "jpeg_quality": 80,
            })

        self.assertTrue(result.success, result.message)
        self.assertEqual(
            [Path(output).name for output in result.outputs],
            ["document_page_002.jpg", "document_page_001.jpg", "document_page_002(2).jpg"],
        )

    def test_last_image_uses_same_page_size_as_main_document(self):
        source = self.folder / "source.pdf"
        backend.images_to_pdf([str(self.png)], source)
        output = self.folder / "appended.pdf"
        backend.append_part(source, self.jpg, output, None)
        with pymupdf.open(output) as document:
            self.assertEqual(document.page_count, 2)
            self.assertAlmostEqual(document[0].rect.width, document[1].rect.width)
            self.assertAlmostEqual(document[0].rect.height, document[1].rect.height)

    def test_resize_creates_new_names_and_keeps_proportions(self):
        outputs = backend.resize_images([str(self.png), str(self.jpg)], 50, 85)
        self.assertEqual(Path(outputs[0]).name, "one_resized_50.png")
        with Image.open(outputs[0]) as image:
            self.assertEqual(image.size, (60, 40))
        with Image.open(outputs[1]) as image:
            self.assertEqual(image.size, (40, 60))

    def test_lossy_and_lossless_image_compression(self):
        for lossy in (False, True):
            outputs = backend.compress_images(
                [str(self.png), str(self.jpg)], lossy, 75, 64
            )
            for output in outputs:
                with Image.open(output) as image:
                    image.verify()


class PdfCompressionSelectionTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.folder = Path(self.temporary.name)
        self.source = self.folder / "source.pdf"
        self.source.write_bytes(b"x" * 1000)

    def tearDown(self):
        self.temporary.cleanup()

    @staticmethod
    def writer(size):
        return lambda _source, output, *args: output.write_bytes(b"x" * size)

    def test_twenty_percent_stops_additional_candidate(self):
        with mock.patch.object(backend, "compress_lossless", side_effect=self.writer(750)), \
             mock.patch.object(backend, "compress_lossy", side_effect=self.writer(800)), \
             mock.patch.object(backend, "compress_with_pikepdf") as pike:
            result = backend.compress_pdf_best(self.source, self.folder, True)
        self.assertEqual(result.suffix, "_pymu")
        pike.assert_not_called()

    def test_smaller_candidate_is_kept_below_twenty_percent(self):
        def lossless(source, output):
            output.write_bytes(b"x" * (900 if source == self.source else 880))
        with mock.patch.object(backend, "compress_lossless", side_effect=lossless), \
             mock.patch.object(backend, "compress_lossy", side_effect=self.writer(920)), \
             mock.patch.object(backend, "compress_with_pikepdf", side_effect=self.writer(890)):
            result = backend.compress_pdf_best(self.source, self.folder, True)
        self.assertTrue(result.reduced)
        self.assertEqual(result.suffix, "_gs_pymu")

    def test_batch_compression_keeps_successful_results(self):
        second = self.folder / "second.pdf"
        second.write_bytes(b"y" * 1000)

        def compression(source, _work, _extra, _profile):
            reduced = source.name == self.source.name
            return backend.CompressionResult(
                source, "_pymu" if reduced else "", reduced, False,
            )

        with mock.patch.object(
            backend, "compress_pdf_best", side_effect=compression,
        ), mock.patch.object(backend, "atomic_pdf"):
            result = backend.process_job({
                "mode": "compress_mode",
                "pdf_paths": [str(self.source), str(second)],
                "profile": "balanced",
            })

        self.assertTrue(result.success)
        self.assertEqual(
            [Path(path).name for path in result.outputs],
            ["source_pymu.pdf"],
        )
        self.assertEqual(
            [notice.code for notice in result.notices],
            ["some_pdfs_not_reduced"],
        )

    def test_batch_output_folder_disambiguates_equal_names(self):
        first_folder = self.folder / "first"
        second_folder = self.folder / "second"
        output_folder = self.folder / "output"
        first_folder.mkdir()
        second_folder.mkdir()
        output_folder.mkdir()
        sources = [first_folder / "document.pdf", second_folder / "document.pdf"]
        for source in sources:
            source.write_bytes(b"x" * 1000)

        def compression(source, _work, _extra, _profile):
            return backend.CompressionResult(source, "_pymu", True, False)

        with mock.patch.object(
            backend, "compress_pdf_best", side_effect=compression,
        ), mock.patch.object(backend, "atomic_pdf"):
            result = backend.process_job({
                "mode": "compress_mode",
                "pdf_paths": [str(source) for source in sources],
                "output_folder": str(output_folder),
            })

        self.assertEqual(
            [Path(path).name for path in result.outputs],
            ["document_pymu.pdf", "document_pymu(2).pdf"],
        )


if __name__ == "__main__":
    unittest.main()
