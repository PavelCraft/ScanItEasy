from __future__ import annotations

import json
import logging
import os
import platform
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Callable, Iterable

import pymupdf
from PIL import Image, ImageOps

LOGGER = logging.getLogger("ScanItEasy")


class ScanItEasyError(Exception):
    def __init__(self, code: str, **params) -> None:
        super().__init__(code)
        self.code = code
        self.params = params


class InputFileOpenError(ScanItEasyError):
    pass


class OutputFileLockedError(ScanItEasyError):
    pass


class CompressionNoGainError(ScanItEasyError):
    def __init__(self, ghostscript_missing: bool = False) -> None:
        super().__init__("compression_no_gain")
        self.ghostscript_missing = ghostscript_missing


@dataclass
class Notice:
    code: str
    params: dict = field(default_factory=dict)


@dataclass
class JobResult:
    success: bool
    outputs: list[str] = field(default_factory=list)
    message: str = ""
    notices: list[Notice] = field(default_factory=list)
    error_code: str | None = None
    ghostscript_missing: bool = False
    message_params: dict = field(default_factory=dict)


@dataclass
class CompressionResult:
    path: Path
    suffix: str
    reduced: bool
    ghostscript_missing: bool = False


@dataclass
class OpenResult:
    opened: bool
    used_acrobat: bool = False
    acrobat_missing: bool = False
    default_app_missing: bool = False


def resource_path(relative_path: str) -> Path:
    import sys
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / relative_path


def app_data_dir() -> Path:
    path = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ScanItEasy"
    path.mkdir(parents=True, exist_ok=True)
    return path


def load_settings() -> dict:
    try:
        return json.loads((app_data_dir() / "settings.json").read_text("utf-8"))
    except (OSError, ValueError, TypeError):
        return {}


def save_settings(values: dict) -> None:
    target = app_data_dir() / "settings.json"
    temporary = target.with_suffix(".tmp")
    temporary.write_text(json.dumps(values, ensure_ascii=False, indent=2), "utf-8")
    os.replace(temporary, target)


def parse_page_sequence(text: str) -> list[int]:
    """Return 1-based pages, preserving order, reverse ranges and duplicates."""
    if not text or not text.strip():
        return []
    result: list[int] = []
    for raw in text.split(","):
        part = raw.strip()
        if not part:
            raise ValueError("empty_page_item")
        if "-" in part:
            pieces = [item.strip() for item in part.split("-")]
            if len(pieces) != 2 or not all(item.isdigit() for item in pieces):
                raise ValueError("invalid_page_range")
            start, end = map(int, pieces)
            if min(start, end) < 1:
                raise ValueError("page_numbers_start_at_one")
            step = 1 if end >= start else -1
            result.extend(range(start, end + step, step))
        elif part.isdigit() and int(part) >= 1:
            result.append(int(part))
        else:
            raise ValueError("page_numbers_start_at_one")
    return result


def parse_and_adjust_indices(text: str) -> list[int]:
    return [page - 1 for page in parse_page_sequence(text)]


def select_pages(requested: str | list[int] | None, count: int) -> tuple[list[int], list[int]]:
    if requested in (None, "", "all"):
        return list(range(count)), []
    pages = parse_page_sequence(requested) if isinstance(requested, str) else [p + 1 for p in requested]
    selected = [page - 1 for page in pages if 1 <= page <= count]
    skipped = [page for page in pages if not 1 <= page <= count]
    if not selected:
        raise ScanItEasyError("no_requested_pages")
    return selected, skipped


def ensure_replaceable(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        try:
            with path.open("r+b"):
                pass
        except OSError as error:
            raise OutputFileLockedError("output_locked", filename=path.name) from error


def validate_pdf(path: Path) -> None:
    try:
        with pymupdf.open(path) as document:
            if path.stat().st_size == 0 or document.page_count == 0:
                raise ScanItEasyError("empty_pdf")
    except ScanItEasyError:
        raise
    except Exception as error:
        raise ScanItEasyError("invalid_pdf") from error


def atomic_pdf(source: Path, destination: Path) -> None:
    validate_pdf(source)
    ensure_replaceable(destination)
    os.replace(source, destination)


def is_docx_open_in_word(path: Path) -> bool:
    if platform.system() != "Windows":
        return False
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return False
        target = path.resolve()
        for document in word.Documents:
            try:
                if Path(document.FullName).resolve() == target:
                    return True
            except Exception:
                continue
        return False
    finally:
        pythoncom.CoUninitialize()


def convert_docx(path: Path, output: Path) -> None:
    if is_docx_open_in_word(path):
        raise InputFileOpenError("word_document_open")
    if platform.system() != "Windows":
        from docx2pdf import convert
        convert(str(path), str(output), keep_active=True)
        return
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(path.resolve()), ReadOnly=True)
        document.ExportAsFixedFormat(str(output.resolve()), 17)
    except Exception as error:
        raise ScanItEasyError("word_conversion_failed") from error
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        pythoncom.CoUninitialize()


def copy_pages(source: Path, output: Path, requested: str | None) -> list[int]:
    with pymupdf.open(source) as original, pymupdf.open() as result:
        pages, skipped = select_pages(requested, original.page_count)
        for page in pages:
            result.insert_pdf(original, from_page=page, to_page=page)
        result.save(output, garbage=4, deflate=True, use_objstms=1)
    return skipped


def raster_pdf(source: Path, output: Path, requested: str | None, dpi: int, gray: bool) -> list[int]:
    """Rasterize directly to PDF, avoiding intermediate PNG files on disk."""
    with pymupdf.open(source) as original, pymupdf.open() as result:
        pages, skipped = select_pages(requested, original.page_count)
        matrix = pymupdf.Matrix(dpi / 72, dpi / 72)
        colorspace = pymupdf.csGRAY if gray else pymupdf.csRGB
        for number in pages:
            source_page = original[number]
            pixmap = source_page.get_pixmap(matrix=matrix, colorspace=colorspace, alpha=False)
            page = result.new_page(width=source_page.rect.width, height=source_page.rect.height)
            page.insert_image(page.rect, stream=pixmap.tobytes("png"))
        result.save(output, garbage=4, deflate=True, use_objstms=1)
    return skipped


def decorate_pdf(source: Path, output: Path, color: str) -> None:
    data = {
        "red": ("Красная лента. Первая страница.png", "Уголок с красной лентой.png", (83.3, 44.88), (91.5, 136.8)),
        "blue": ("Синяя лента. Первая страница.png", "Уголок с синей лентой.png", (110, 110), (110, 110)),
    }
    first_name, other_name, first_size, other_size = data[color]
    first = resource_path(f"doc_decorations/{first_name}")
    other = resource_path(f"doc_decorations/{other_name}")
    with pymupdf.open(source) as document:
        if document.page_count:
            document[0].insert_image(pymupdf.Rect(0, 0, *first_size), filename=str(first))
        for number in range(1, document.page_count):
            width, height = other_size
            document[number].insert_image(pymupdf.Rect(0, 0, width, height), filename=str(other))
        document.save(output, garbage=4, deflate=True, use_objstms=1)


def _oriented_bytes(path: Path) -> tuple[bytes, int, int]:
    with Image.open(path) as original:
        orientation = original.getexif().get(274, 1)
        image = ImageOps.exif_transpose(original)
        width, height = image.size
        if path.suffix.lower() in {".jpg", ".jpeg"} and orientation == 1:
            return path.read_bytes(), width, height
        buffer = BytesIO()
        if path.suffix.lower() in {".jpg", ".jpeg"}:
            image.convert("RGB").save(buffer, "JPEG", quality=95, optimize=True)
        else:
            image.save(buffer, "PNG")
        return buffer.getvalue(), width, height


def add_image_page(
    document: pymupdf.Document,
    path: Path,
    page_size: tuple[float, float] | None = None,
) -> None:
    data, width, height = _oriented_bytes(path)
    if page_size is None:
        page_width, page_height = width * 72 / 96, height * 72 / 96
    else:
        page_width, page_height = page_size
    page = document.new_page(width=page_width, height=page_height)
    scale = min(page_width / width, page_height / height)
    image_width, image_height = width * scale, height * scale
    left = (page_width - image_width) / 2
    top = (page_height - image_height) / 2
    target = pymupdf.Rect(left, top, left + image_width, top + image_height)
    page.insert_image(target, stream=data)


def append_part(source: Path, tail_path: Path, output: Path, requested: str | None) -> list[int]:
    skipped: list[int] = []
    with pymupdf.open(source) as document:
        if tail_path.suffix.lower() == ".pdf":
            with pymupdf.open(tail_path) as tail:
                pages, skipped = select_pages(requested, tail.page_count)
                for page in pages:
                    document.insert_pdf(tail, from_page=page, to_page=page)
        else:
            reference = document[-1].rect if document.page_count else pymupdf.paper_rect("a4")
            add_image_page(document, tail_path, (reference.width, reference.height))
        document.save(output, garbage=4, deflate=True, use_objstms=1)
    return skipped


def compress_lossless(source: Path, output: Path) -> None:
    with pymupdf.open(source) as document:
        document.save(output, garbage=4, deflate=True, use_objstms=1)


def compress_with_pikepdf(source: Path, output: Path) -> None:
    import pikepdf
    with pikepdf.open(source) as document:
        document.save(
            output,
            compress_streams=True,
            recompress_flate=True,
            object_stream_mode=pikepdf.ObjectStreamMode.generate,
        )
    validate_pdf(output)


def compress_lossy(source: Path, output: Path, profile: str) -> None:
    from pdf_compressor import get_ghostscript_path
    dpi = {"high": 300, "balanced": 180, "small": 100}.get(profile, 180)
    command = [get_ghostscript_path(), "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.6",
               "-dNOPAUSE", "-dQUIET", "-dBATCH", "-dDetectDuplicateImages=true",
               "-dDownsampleColorImages=true", "-dDownsampleGrayImages=true",
               f"-dColorImageResolution={dpi}", f"-dGrayImageResolution={dpi}",
               f"-sOutputFile={output}", str(source)]
    subprocess.run(command, check=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    validate_pdf(output)


def _reduction(source: Path, candidate: Path) -> float:
    return 1 - candidate.stat().st_size / source.stat().st_size


def compress_pdf_best(source: Path, work: Path, extra: bool,
                      profile: str = "balanced") -> CompressionResult:
    """Build only useful candidates and return the smallest valid PDF.

    Twenty percent is a search threshold, never an acceptance threshold.
    """
    candidates: list[tuple[Path, str]] = [(source, "")]
    pymu = work / "pymu.pdf"
    compress_lossless(source, pymu)
    candidates.append((pymu, "_pymu"))
    ghostscript_missing = False

    if extra:
        ghostscript = work / "gs.pdf"
        try:
            compress_lossy(source, ghostscript, profile)
            candidates.append((ghostscript, "_gs"))
            gs_pymu = work / "gs_pymu.pdf"
            compress_lossless(ghostscript, gs_pymu)
            candidates.append((gs_pymu, "_gs_pymu"))
            if max(_reduction(source, path) for path, _ in candidates) < 0.20:
                gs_pike = work / "gs_pike.pdf"
                compress_with_pikepdf(ghostscript, gs_pike)
                candidates.append((gs_pike, "_gs_pike"))
        except FileNotFoundError:
            ghostscript_missing = True
            if _reduction(source, pymu) < 0.20:
                pike = work / "pike.pdf"
                compress_with_pikepdf(source, pike)
                candidates.append((pike, "_pike"))
        except (subprocess.CalledProcessError, OSError):
            if _reduction(source, pymu) < 0.20:
                pike = work / "pike.pdf"
                compress_with_pikepdf(source, pike)
                candidates.append((pike, "_pike"))

    winner, suffix = min(candidates, key=lambda item: item[0].stat().st_size)
    reduced = winner.stat().st_size < source.stat().st_size
    if not reduced:
        winner, suffix = source, ""
    return CompressionResult(winner, suffix, reduced, ghostscript_missing)


def merge_multiple_pdfs(files: list[str], pages: list[str | list[int]], output_path: str) -> list[int]:
    destination = Path(output_path).resolve()
    ensure_replaceable(destination)
    skipped_all: list[int] = []
    with tempfile.TemporaryDirectory(dir=destination.parent, prefix=".sie-") as folder:
        temporary = Path(folder) / "merged.pdf"
        with pymupdf.open() as result:
            for filename, requested in zip(files, pages):
                with pymupdf.open(filename) as source:
                    selected, skipped = select_pages(requested, source.page_count)
                    skipped_all.extend(skipped)
                    for page in selected:
                        result.insert_pdf(source, from_page=page, to_page=page)
            result.save(temporary, garbage=4, deflate=True, use_objstms=1)
        atomic_pdf(temporary, destination)
    return skipped_all


def pdf_to_images(source: Path, kind: str, requested: str | None, dpi: int,
                  quality: int, output_base: Path | None = None) -> tuple[list[str], list[int]]:
    outputs: list[str] = []
    suffix = ".jpg" if kind == "jpeg" else ".png"
    output_base = output_base or source
    with pymupdf.open(source) as document:
        pages, skipped = select_pages(requested, document.page_count)
        digits = max(3, len(str(document.page_count)))
        occurrences: dict[int, int] = {}
        planned: list[Path] = []
        for page_index in pages:
            page_number = page_index + 1
            occurrences[page_number] = occurrences.get(page_number, 0) + 1
            duplicate = f"({occurrences[page_number]})" if occurrences[page_number] > 1 else ""
            planned.append(output_base.with_name(
                f"{output_base.stem}_page_{page_number:0{digits}d}{duplicate}{suffix}"
            ))
        for output in planned:
            ensure_replaceable(output)
        for index, page_number in enumerate(pages, 1):
            output = planned[index - 1]
            pixmap = document[page_number].get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False)
            if kind == "jpeg":
                with Image.open(BytesIO(pixmap.tobytes("png"))) as image:
                    image.convert("RGB").save(output, "JPEG", quality=quality, optimize=True, progressive=True)
            else:
                pixmap.save(output)
            outputs.append(str(output))
    return outputs, skipped


def images_to_pdf(files: list[str], output: Path) -> None:
    if not files:
        raise ScanItEasyError("no_images")
    ensure_replaceable(output)
    with tempfile.TemporaryDirectory(dir=output.parent, prefix=".sie-") as folder:
        temporary = Path(folder) / "images.pdf"
        with pymupdf.open() as document:
            for filename in files:
                add_image_page(document, Path(filename))
            document.save(temporary, garbage=4, deflate=True, use_objstms=1)
        atomic_pdf(temporary, output)


def compress_images(files: Iterable[str], lossy: bool, quality: int, colors: int) -> list[str]:
    files = list(files)
    outputs: list[str] = []
    for filename in files:
        source = Path(filename)
        detail = f"jpg{quality}" if source.suffix.lower() in {".jpg", ".jpeg"} and lossy else (
            f"png{colors}" if lossy else "lossless")
        output = source.with_name(f"{source.stem}_pillow_{detail}{source.suffix.lower()}")
        ensure_replaceable(output)
        with tempfile.TemporaryDirectory(dir=source.parent, prefix=".sie-") as folder:
            candidate = Path(folder) / output.name
            with Image.open(source) as original:
                image = ImageOps.exif_transpose(original)
                if source.suffix.lower() in {".jpg", ".jpeg"}:
                    if not lossy:
                        original.save(candidate, "JPEG", quality="keep", optimize=True,
                                      progressive=True, exif=original.getexif())
                    else:
                        image.convert("RGB").save(candidate, "JPEG", quality=quality,
                                                  optimize=True, progressive=True)
                elif lossy:
                    image.convert("RGBA").quantize(colors=colors).save(candidate, "PNG", optimize=True)
                else:
                    image.save(candidate, "PNG", optimize=True, compress_level=9)
            with Image.open(candidate) as check:
                check.verify()
            if candidate.stat().st_size < source.stat().st_size:
                os.replace(candidate, output)
                outputs.append(str(output))
    return outputs


def resize_images(files: Iterable[str], percent: int, quality: int) -> list[str]:
    if not 1 <= percent <= 500:
        raise ScanItEasyError("invalid_resize_percent")
    files = list(files)
    planned = [Path(name).with_name(
        f"{Path(name).stem}_resized_{percent}{Path(name).suffix.lower()}") for name in files]
    for output in planned:
        ensure_replaceable(output)
    outputs: list[str] = []
    for filename, output in zip(files, planned):
        source = Path(filename)
        with Image.open(source) as original:
            image = ImageOps.exif_transpose(original)
            size = (max(1, round(image.width * percent / 100)), max(1, round(image.height * percent / 100)))
            resized = image.resize(size, Image.Resampling.LANCZOS)
            if source.suffix.lower() in {".jpg", ".jpeg"}:
                resized.convert("RGB").save(output, "JPEG", quality=quality, optimize=True, progressive=True)
            else:
                resized.save(output, "PNG", optimize=True, compress_level=9)
        outputs.append(str(output))
    return outputs


def _skipped_notice(skipped: list[int]) -> Notice | None:
    return Notice("missing_pages", {"pages": ", ".join(map(str, skipped))}) if skipped else None


def process_job(settings: dict, progress: Callable[[int], None] | None = None) -> JobResult:
    report = progress or (lambda _value: None)
    notices: list[Notice] = []
    ghostscript_missing = False
    mode = settings.get("mode")
    report(5)
    try:
        if mode in {"default_mode", "custom_mode", "convert_to_pdf_mode"}:
            source = Path(settings["docx_path"]).resolve()
            destination = source.with_suffix(".pdf")
            if is_docx_open_in_word(source):
                raise InputFileOpenError("word_document_open")
            with tempfile.TemporaryDirectory(dir=destination.parent, prefix=".sie-") as folder:
                work = Path(folder)
                converted = work / "converted.pdf"
                convert_docx(source, converted)
                report(25)
                current = work / "selected.pdf"
                notice = _skipped_notice(copy_pages(converted, current, settings.get("pages")))
                if notice: notices.append(notice)
                if mode != "convert_to_pdf_mode":
                    raster = work / "raster.pdf"
                    raster_pdf(current, raster, None, 300, True)
                    decorated = work / "decorated.pdf"
                    decorate_pdf(raster, decorated, settings.get("color", "red"))
                    current = decorated
                    last_part = settings.get("last_page_path")
                    if last_part:
                        appended = work / "appended.pdf"
                        notice = _skipped_notice(append_part(current, Path(last_part), appended, settings.get("last_page_pages")))
                        if notice: notices.append(notice)
                        current = appended
                    report(65)
                compression = compress_pdf_best(
                    current, work, bool(settings.get("is_compression_needed")),
                    settings.get("profile", "balanced"),
                )
                current = compression.path
                ghostscript_missing = compression.ghostscript_missing
                destination = source.with_name(f"{source.stem}{compression.suffix}.pdf")
                ensure_replaceable(destination)
                if settings.get("is_compression_needed") and not compression.reduced:
                    notices.append(Notice("compression_no_extra_gain"))
                result = work / "result.pdf"
                shutil.copyfile(current, result)
                atomic_pdf(result, destination)
            outputs = [str(destination)]
        elif mode == "merge_pdf":
            files, pages, output = settings["parameters"]
            requested = Path(output).resolve()
            with tempfile.TemporaryDirectory(dir=requested.parent, prefix=".sie-") as folder:
                work = Path(folder)
                merged = work / "merged.pdf"
                notice = _skipped_notice(merge_multiple_pdfs(files, pages, str(merged)))
                if notice: notices.append(notice)
                compression = compress_pdf_best(
                    merged, work, bool(settings.get("is_compression_needed")),
                    settings.get("profile", "balanced"),
                )
                ghostscript_missing = compression.ghostscript_missing
                if settings.get("is_compression_needed") and not compression.reduced:
                    notices.append(Notice("compression_no_extra_gain"))
                destination = requested.with_name(
                    f"{requested.stem}{compression.suffix}.pdf"
                )
                atomic_pdf(compression.path, destination)
            outputs = [str(destination)]
        elif mode == "grayscale_mode":
            source = Path(settings["pdf_path"]).resolve()
            requested = source.with_name(f"{source.stem}_grayscale.pdf")
            with tempfile.TemporaryDirectory(dir=requested.parent, prefix=".sie-") as folder:
                work = Path(folder)
                temporary = work / "gray.pdf"
                notice = _skipped_notice(raster_pdf(source, temporary, settings.get("pages"), 300, True))
                if notice: notices.append(notice)
                compression = compress_pdf_best(
                    temporary, work, bool(settings.get("is_compression_needed")),
                    settings.get("profile", "balanced"),
                )
                ghostscript_missing = compression.ghostscript_missing
                if settings.get("is_compression_needed") and not compression.reduced:
                    notices.append(Notice("compression_no_extra_gain"))
                destination = requested.with_name(
                    f"{requested.stem}{compression.suffix}.pdf"
                )
                atomic_pdf(compression.path, destination)
            outputs = [str(destination)]
        elif mode == "compress_mode":
            source = Path(settings["pdf_path"]).resolve()
            with tempfile.TemporaryDirectory(dir=source.parent, prefix=".sie-") as folder:
                compression = compress_pdf_best(
                    source, Path(folder), True, settings.get("profile", "balanced")
                )
                ghostscript_missing = compression.ghostscript_missing
                if not compression.reduced:
                    raise CompressionNoGainError(ghostscript_missing)
                destination = source.with_name(f"{source.stem}{compression.suffix}.pdf")
                ensure_replaceable(destination)
                atomic_pdf(compression.path, destination)
            outputs = [str(destination)]
        elif mode in {"convert_to_png_mode", "convert_to_jpg_mode"}:
            kind = "jpeg" if mode == "convert_to_jpg_mode" else "png"
            outputs, skipped = pdf_to_images(Path(settings["pdf_path"]), kind, settings.get("pages"), int(settings.get("dpi", 200)), int(settings.get("jpeg_quality", 85)))
            notice = _skipped_notice(skipped)
            if notice: notices.append(notice)
        elif mode in {"convert_docx_to_png_mode", "convert_docx_to_jpg_mode"}:
            source = Path(settings["docx_path"]).resolve()
            if is_docx_open_in_word(source):
                raise InputFileOpenError("word_document_open")
            with tempfile.TemporaryDirectory(dir=source.parent, prefix=".sie-") as folder:
                converted = Path(folder) / "converted.pdf"
                convert_docx(source, converted)
                report(45)
                kind = "jpeg" if mode == "convert_docx_to_jpg_mode" else "png"
                outputs, skipped = pdf_to_images(
                    converted, kind, settings.get("pages"),
                    int(settings.get("dpi", 200)), int(settings.get("jpeg_quality", 85)),
                    output_base=source,
                )
            notice = _skipped_notice(skipped)
            if notice: notices.append(notice)
        elif mode == "images_to_pdf_mode":
            files = settings.get("image_paths", [])
            requested = Path(files[0]).with_name(f"{Path(files[0]).stem}_images.pdf")
            with tempfile.TemporaryDirectory(dir=requested.parent, prefix=".sie-") as folder:
                work = Path(folder)
                temporary = work / "images.pdf"
                images_to_pdf(files, temporary)
                compression = compress_pdf_best(
                    temporary, work, bool(settings.get("is_compression_needed")),
                    settings.get("profile", "balanced"),
                )
                ghostscript_missing = compression.ghostscript_missing
                if settings.get("is_compression_needed") and not compression.reduced:
                    notices.append(Notice("compression_no_extra_gain"))
                output = requested.with_name(
                    f"{requested.stem}{compression.suffix}.pdf"
                )
                atomic_pdf(compression.path, output)
            outputs = [str(output)]
        elif mode == "compress_images_mode":
            image_paths = settings.get("image_paths", [])
            outputs = compress_images(image_paths, bool(settings.get("lossy")), int(settings.get("jpeg_quality", 85)), int(settings.get("png_colors", 256)))
            if not outputs:
                raise CompressionNoGainError()
            if len(outputs) < len(image_paths):
                notices.append(Notice("some_images_not_reduced"))
        elif mode == "resize_images_mode":
            outputs = resize_images(settings.get("image_paths", []), int(settings.get("resize_percent", 100)), int(settings.get("jpeg_quality", 85)))
        else:
            raise ScanItEasyError("unknown_mode")
        report(100)
        return JobResult(True, outputs, "done", notices,
                         ghostscript_missing=ghostscript_missing)
    except InputFileOpenError as error:
        return JobResult(False, message=error.code, error_code="input_open",
                         message_params=error.params)
    except OutputFileLockedError as error:
        return JobResult(False, message=error.code, error_code="output_locked",
                         message_params=error.params)
    except CompressionNoGainError as error:
        return JobResult(False, message=error.code, error_code="compression_no_gain",
                         ghostscript_missing=error.ghostscript_missing)
    except ScanItEasyError as error:
        return JobResult(False, message=error.code, error_code="processing",
                         message_params=error.params)
    except Exception as error:
        LOGGER.exception("Job failed")
        return JobResult(False, message="processing_failed", error_code="processing",
                         message_params={"error": str(error)})


def work_process(settings: dict) -> None:
    settings["result"] = process_job(settings, lambda value: settings.__setitem__("progress", value))


def find_acrobat(saved: str | None = None) -> Path | None:
    candidates = [Path(saved)] if saved else []
    if platform.system() == "Windows":
        import winreg
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe") as key:
                    candidates.append(Path(winreg.QueryValue(key, None).strip('"')))
            except OSError:
                pass
        for name in ("ProgramFiles", "ProgramFiles(x86)"):
            if os.environ.get(name):
                candidates.append(Path(os.environ[name]) / "Adobe/Acrobat DC/Acrobat/Acrobat.exe")
    for candidate in candidates:
        if candidate.is_absolute() and candidate.is_file() and candidate.name.lower() == "acrobat.exe":
            return candidate.resolve()
    return None


def acrobat_from_folder(folder: str) -> Path | None:
    directory = Path(folder)
    direct = directory / "Acrobat.exe"
    if direct.is_file():
        return direct.resolve()
    return next((path.resolve() for path in directory.glob("*/Acrobat.exe") if path.is_file()), None)


def open_default(path: Path) -> bool:
    try:
        if platform.system() == "Windows": os.startfile(str(path))
        elif platform.system() == "Darwin": subprocess.Popen(["open", str(path)])
        else: subprocess.Popen(["xdg-open", str(path)])
        return True
    except (OSError, FileNotFoundError):
        return False


def open_result(path: str, prefer_acrobat: bool, requested_page: str | int, saved_path: str | None) -> OpenResult:
    pdf = Path(path).resolve()
    try: requested = int(requested_page)
    except (TypeError, ValueError): requested = 1
    with pymupdf.open(pdf) as document:
        page = min(max(1, requested), max(1, document.page_count))
    if prefer_acrobat:
        acrobat = find_acrobat(saved_path)
        if acrobat:
            try:
                subprocess.Popen([str(acrobat), "/A", f"page={page}", str(pdf)], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                return OpenResult(True, used_acrobat=True)
            except OSError:
                pass
        opened = open_default(pdf)
        return OpenResult(opened, acrobat_missing=True, default_app_missing=not opened)
    opened = open_default(pdf)
    return OpenResult(opened, default_app_missing=not opened)
