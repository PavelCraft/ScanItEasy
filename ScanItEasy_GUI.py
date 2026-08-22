from __future__ import annotations

import logging
import locale
import os
import subprocess
import threading
import tkinter as tk
import textwrap
import webbrowser
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk
from PIL import Image, ImageDraw

from ScanItEasy_backend import (
    JobResult,
    acrobat_from_folder,
    load_settings,
    open_result,
    parse_page_sequence,
    process_job,
    resource_path,
    save_settings,
)


GREEN = "#2FA572"
RED = "#C62828"
DARK_RED = "#8B0000"
TEXT = ("gray10", "gray90")
APP_VERSION = "2.1.0"
GUIDE_URLS = {
    "ru": "https://pavelcraft.github.io/ScanItEasy/",
    "en": "https://pavelcraft.github.io/ScanItEasy/en/",
}
GHOSTSCRIPT_GUIDE_URLS = {
    "ru": "https://pavelcraft.github.io/ScanItEasy/ghostscript/",
    "en": "https://pavelcraft.github.io/ScanItEasy/en/ghostscript/",
}
EMAIL_URL = "mailto:contact@formacode.dev?subject=ScanItEasy%20feedback"
TELEGRAM_URL = "https://t.me/PavelCraftRobot"

MODES = [
    ("custom_mode", "mode_scan"),
    ("convert_to_pdf_mode", "mode_docx_pdf"),
    ("convert_docx_to_png_mode", "mode_docx_png"),
    ("convert_docx_to_jpg_mode", "mode_docx_jpg"),
    ("merge_pdf_mode", "mode_merge"),
    ("grayscale_mode", "mode_grayscale"),
    ("compress_mode", "mode_compress_pdf"),
    ("convert_to_png_mode", "mode_pdf_png"),
    ("convert_to_jpg_mode", "mode_pdf_jpg"),
    ("images_to_pdf_mode", "mode_images_pdf"),
    ("compress_images_mode", "mode_compress_images"),
    ("resize_images_mode", "mode_resize_images"),
]

STRINGS = {
    "ru": {
        "choose_mode": "Выберите режим:", "main_mode": "Основной режим",
        "additional": "Дополнительно:", "about": "О приложении",
        "mode_scan": "Имитация сканирования", "mode_docx_pdf": "Конвертировать DOCX в PDF",
        "mode_docx_png": "Конвертировать DOCX в PNG", "mode_docx_jpg": "Конвертировать DOCX в JPG",
        "mode_merge": "Объединить PDF", "mode_grayscale": "Сделать PDF чёрно-белым",
        "mode_compress_pdf": "Сжать PDF", "mode_pdf_png": "Конвертировать PDF в PNG",
        "mode_pdf_jpg": "Конвертировать PDF в JPG", "mode_images_pdf": "Изображения в PDF",
        "mode_compress_images": "Сжать изображения", "mode_resize_images": "Изменить размер изображений",
        "select_docx": "Выберите DOCX-файл", "select_pdf": "Выберите PDF-файл",
        "select_pages": "Выберите страницы", "ribbon_color": "Цвет ленточки",
        "red": "Красный", "blue": "Синий", "select_last_part": "Выберите последнюю часть документа",
        "last_part_pages": "Страницы последней части", "resolution_dpi": "Разрешение, DPI",
        "jpg_quality": "Качество JPG, %", "quality": "Качество",
        "quality_high": "Высокое (300 DPI)", "quality_balanced": "Сбалансированное (180 DPI)",
        "quality_small": "Минимальный размер (100 DPI)", "lossy_images": "Сжать с потерей качества",
        "png_colors": "Количество цветов PNG", "resize_percent": "Изменить размер, %",
        "jpg_save_quality": "Качество сохранения JPG, %", "compress_pdf": "Сжать PDF",
        "scan": "Сканировать", "compress": "Сжать", "change": "Изменить",
        "merge": "Объединить", "convert": "Конвертировать", "open_folder": "Открыть папку",
        "continue": "Продолжить", "specify_path": "Указать путь", "choose_file": "Выбрать файл",
        "all_pages_hint": "Если нужны все страницы, оставьте поле пустым",
        "pages_placeholder": "Например: 7, 1-5 или 5-1",
        "select_images": "Выберите одно или несколько изображений", "choose_images": "Выбрать изображения",
        "select_pdfs": "Выберите два или более PDF-файлов", "add_files": "Добавить файлы",
        "select_pdfs_to_compress": "Выберите один или несколько PDF-файлов",
        "compression_different_folders": "Файлы выбраны из разных папок. При желании сохраните все результаты в одну папку.",
        "choose_output_folder": "Выбрать папку для результатов",
        "merge_pages_placeholder": "Страницы: 7, 1-5 или 5-1; пусто — все",
        "open_result": "Открыть итоговый файл", "open_default": "В приложении по умолчанию",
        "open_acrobat": "В Adobe Acrobat на странице №", "images_type": "Изображения",
        "word_docs_type": "Документы Word", "pdf_or_image_type": "PDF или изображение",
        "choose_docx_error": "Выберите DOCX-файл.", "choose_two_pdfs_error": "Выберите хотя бы два PDF-файла.",
        "check_pages": "Проверьте номера страниц.", "choose_pdf_error": "Выберите PDF-файл.",
        "choose_image_error": "Выберите хотя бы одно изображение.",
        "check_parameters": "Проверьте номера страниц и числовые параметры.",
        "processing": "Выполняется обработка…", "no_worker_result": "Рабочий процесс завершился без результата.",
        "done": "Готово.", "compression_no_gain": "Уменьшить файл не удалось.",
        "compression_no_extra_gain": "Дополнительное сжатие не уменьшило файл.",
        "some_images_not_reduced": "Некоторые изображения уже были меньше полученных вариантов; для них новые файлы не созданы.",
        "some_pdfs_not_reduced": "Некоторые PDF не удалось уменьшить; для них новые файлы не созданы.",
        "missing_pages": "Пропущены отсутствующие страницы: {pages}.",
        "word_document_open": "Документ открыт в Microsoft Word. Сохраните и закройте его, затем нажмите «Продолжить».",
        "output_locked": "Невозможно заменить файл «{filename}». Закройте его и нажмите «Продолжить».",
        "no_requested_pages": "Ни одной из указанных страниц нет в документе.",
        "empty_pdf": "Созданный PDF не содержит страниц.", "invalid_pdf": "Созданный PDF повреждён или не был создан.",
        "word_conversion_failed": "Microsoft Word не смог преобразовать документ в PDF.",
        "no_images": "Не выбрано ни одного изображения.", "invalid_resize_percent": "Масштаб должен быть от 1 до 500 процентов.",
        "unknown_mode": "Неизвестный режим работы.", "processing_failed": "Не удалось выполнить операцию: {error}",
        "install_ghostscript": "Установите ", "stronger_compression": "Для более сильного сжатия установите ",
        "acrobat_missing": "Adobe Acrobat не найден.", "opened_default": " Файл открыт в приложении по умолчанию.",
        "no_default_pdf_app": "PDF создан, но в системе нет приложения для его открытия.",
        "acrobat_folder_title": "Укажите папку с Acrobat.exe", "acrobat_path_saved": "Путь к Adobe Acrobat сохранён.",
        "acrobat_not_in_folder": "В выбранной папке Acrobat.exe не найден.",
        "folder_open_failed": "Не удалось открыть папку с результатом.",
        "wait_for_job": "Дождитесь завершения текущей операции.",
        "version": "Версия", "developed_by": "Разработано FormaCode", "contacts": "Контакты",
        "instruction": "Инструкция по использованию", "instruction_hint": "Описание режима имитации сканирования",
        "about_intro": "Создание, преобразование и обработка документов и изображений.",
    },
    "en": {
        "choose_mode": "Choose a mode:", "main_mode": "Main mode", "additional": "Additional:",
        "about": "About", "mode_scan": "Scan simulation", "mode_docx_pdf": "Convert DOCX to PDF",
        "mode_docx_png": "Convert DOCX to PNG", "mode_docx_jpg": "Convert DOCX to JPG",
        "mode_merge": "Merge PDFs", "mode_grayscale": "Convert PDF to grayscale",
        "mode_compress_pdf": "Compress PDF", "mode_pdf_png": "Convert PDF to PNG",
        "mode_pdf_jpg": "Convert PDF to JPG", "mode_images_pdf": "Images to PDF",
        "mode_compress_images": "Compress images", "mode_resize_images": "Resize images",
        "select_docx": "Select a DOCX file", "select_pdf": "Select a PDF file", "select_pages": "Select pages",
        "ribbon_color": "Ribbon color", "red": "Red", "blue": "Blue",
        "select_last_part": "Select the final part of the document", "last_part_pages": "Pages from the final part",
        "resolution_dpi": "Resolution, DPI", "jpg_quality": "JPG quality, %", "quality": "Quality",
        "quality_high": "High (300 DPI)", "quality_balanced": "Balanced (180 DPI)",
        "quality_small": "Smallest size (100 DPI)", "lossy_images": "Allow lossy compression",
        "png_colors": "Number of PNG colors", "resize_percent": "Resize, %",
        "jpg_save_quality": "JPG save quality, %", "compress_pdf": "Compress PDF",
        "scan": "Simulate scan", "compress": "Compress", "change": "Change", "merge": "Merge",
        "convert": "Convert", "open_folder": "Open folder", "continue": "Continue",
        "specify_path": "Specify path", "choose_file": "Choose file",
        "all_pages_hint": "Leave this field empty to use all pages",
        "pages_placeholder": "For example: 7, 1-5 or 5-1",
        "select_images": "Select one or more images", "choose_images": "Choose images",
        "select_pdfs": "Select two or more PDF files", "add_files": "Add files",
        "select_pdfs_to_compress": "Select one or more PDF files",
        "compression_different_folders": "The files are from different folders. You can optionally save all results in one folder.",
        "choose_output_folder": "Choose a results folder",
        "merge_pages_placeholder": "Pages: 7, 1-5 or 5-1; empty means all",
        "open_result": "Open the resulting file", "open_default": "In the default application",
        "open_acrobat": "In Adobe Acrobat on page No.", "images_type": "Images",
        "word_docs_type": "Word documents", "pdf_or_image_type": "PDF or image",
        "choose_docx_error": "Select a DOCX file.", "choose_two_pdfs_error": "Select at least two PDF files.",
        "check_pages": "Check the page numbers.", "choose_pdf_error": "Select a PDF file.",
        "choose_image_error": "Select at least one image.", "check_parameters": "Check page numbers and numeric parameters.",
        "processing": "Processing…", "no_worker_result": "The worker finished without a result.", "done": "Done.",
        "compression_no_gain": "The file could not be reduced.",
        "compression_no_extra_gain": "Additional compression did not reduce the file.",
        "some_images_not_reduced": "Some images were already smaller than the generated variants; no new files were created for them.",
        "some_pdfs_not_reduced": "Some PDFs could not be reduced; no new files were created for them.",
        "missing_pages": "Missing pages were skipped: {pages}.",
        "word_document_open": "The document is open in Microsoft Word. Save and close it, then click Continue.",
        "output_locked": "Cannot replace “{filename}”. Close the file and click Continue.",
        "no_requested_pages": "None of the requested pages exist in the document.",
        "empty_pdf": "The created PDF contains no pages.", "invalid_pdf": "The created PDF is damaged or was not created.",
        "word_conversion_failed": "Microsoft Word could not convert the document to PDF.",
        "no_images": "No images were selected.", "invalid_resize_percent": "Scale must be between 1 and 500 percent.",
        "unknown_mode": "Unknown operation mode.", "processing_failed": "The operation failed: {error}",
        "install_ghostscript": "Install ", "stronger_compression": "For stronger compression, install ",
        "acrobat_missing": "Adobe Acrobat was not found.", "opened_default": " The file was opened in the default application.",
        "no_default_pdf_app": "The PDF was created, but there is no application available to open it.",
        "acrobat_folder_title": "Select the folder containing Acrobat.exe", "acrobat_path_saved": "Adobe Acrobat path saved.",
        "acrobat_not_in_folder": "Acrobat.exe was not found in the selected folder.",
        "folder_open_failed": "Could not open the result folder.", "wait_for_job": "Wait for the current operation to finish.",
        "version": "Version", "developed_by": "Developed by FormaCode", "contacts": "Contacts",
        "instruction": "User guide", "instruction_hint": "How the scan simulation mode works",
        "about_intro": "Create, convert, and process documents and images.",
    },
}


class AutoHideScrollableFrame(ctk.CTkScrollableFrame):
    """Scrollable page whose bar is visible only when content overflows."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self._scrollbar_visible = True
        self.bind("<Configure>", self._schedule_scrollbar_update, add="+")
        self._parent_canvas.bind(
            "<Configure>", self._schedule_scrollbar_update, add="+"
        )
        self.after_idle(self.update_scrollbar_visibility)

    def _schedule_scrollbar_update(self, _event=None) -> None:
        self.after_idle(self.update_scrollbar_visibility)

    def update_scrollbar_visibility(self) -> None:
        if not self.winfo_exists():
            return
        content = self._parent_canvas.bbox("all")
        content_height = 0 if content is None else content[3] - content[1]
        needs_scrollbar = content_height > self._parent_canvas.winfo_height() + 1
        if needs_scrollbar == self._scrollbar_visible:
            return
        self._scrollbar_visible = needs_scrollbar
        if needs_scrollbar:
            self._scrollbar.grid()
        else:
            self._parent_canvas.yview_moveto(0)
            self._scrollbar.grid_remove()


def reveal_file(path: str) -> bool:
    """Open the containing folder and select the newly created file."""
    target = Path(path).resolve()
    try:
        if os.name == "nt":
            # Explorer expects /select, and the path as separate command-line
            # arguments. Combining them makes some Windows versions fall back
            # to the Documents folder, especially for paths containing spaces.
            subprocess.Popen(["explorer.exe", "/select,", str(target)])
        elif __import__("sys").platform == "darwin":
            subprocess.Popen(["open", "-R", str(target)])
        else:
            subprocess.Popen(["xdg-open", str(target.parent)])
        return True
    except OSError:
        return False


def configure_logging() -> None:
    log_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ScanItEasy" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(log_dir / "app.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger = logging.getLogger("ScanItEasy")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        logger.addHandler(handler)


def detect_system_language() -> str:
    try:
        language = locale.getlocale()[0] or ""
    except (ValueError, TypeError):
        language = ""
    return "ru" if language.lower().startswith("ru") else "en"


def is_multipage_pdf(filename: str) -> bool:
    if Path(filename).suffix.lower() != ".pdf":
        return False
    try:
        import pymupdf
        with pymupdf.open(filename) as document:
            return document.page_count > 1
    except (OSError, ValueError, RuntimeError):
        return False


class ScanItEasyApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")
        self.title("ScanItEasy")
        self.geometry("360x450+400+10")
        self.minsize(360, 360)
        try:
            self.iconbitmap(str(resource_path("SIE-no-scan.ico")))
        except tk.TclError:
            pass
        self.preferences = load_settings()
        self.language = self.preferences.get("language")
        if self.language not in STRINGS:
            self.language = detect_system_language()
        self.mode = ""
        self.fields: dict[str, object] = {}
        self.field_frames: dict[str, ctk.CTkFrame] = {}
        self.selected_images: list[str] = []
        self.merge_files: list[str] = []
        self.merge_entries: dict[str, ctk.CTkEntry] = {}
        self.last_job: dict | None = None
        self.progress_value = 0
        self.worker: threading.Thread | None = None
        self.result: JobResult | None = None
        self.protocol("WM_DELETE_WINDOW", self.close_app)
        self.show_menu()

    def tr(self, key: str, **params) -> str:
        template = STRINGS[self.language].get(key, STRINGS["en"].get(key, key))
        try:
            return template.format(**params)
        except (KeyError, ValueError):
            return template

    def file_types(self, kind: str):
        if kind == "images":
            return [(self.tr("images_type"), "*.png *.jpg *.jpeg")]
        if kind == "docx":
            return [(self.tr("word_docs_type"), "*.docx")]
        return [("PDF", "*.pdf")]

    def set_language(self, language: str) -> None:
        if language == self.language:
            return
        self.language = language
        self.preferences["language"] = language
        save_settings(self.preferences)
        self.show_menu()

    def clear(self) -> None:
        for widget in self.winfo_children():
            widget.destroy()
        self.fields.clear()
        self.field_frames.clear()

    def show_menu(self) -> None:
        if "open_result_file" in self.fields:
            self.save_acrobat_preferences()
        self.clear()
        self.mode = ""
        self.minsize(360, 360)
        self.geometry("360x450+400+10")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=0)
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=45, pady=(14, 7), sticky="ew")
        ctk.CTkLabel(header, text=self.tr("choose_mode"), font=("Arial", 16, "bold")).pack(side="left")
        language_frame = ctk.CTkFrame(header, fg_color=("#E5E7EB", "#1B2130"), corner_radius=8)
        language_frame.pack(side="right")
        for code in ("EN", "RU"):
            active = code.lower() == self.language
            ctk.CTkButton(
                language_frame, text=code, width=32, height=24, corner_radius=7,
                font=("Arial", 10, "bold"),
                fg_color=GREEN if active else "transparent",
                hover_color=GREEN, text_color="white" if active else TEXT,
                command=lambda value=code.lower(): self.set_language(value),
            ).pack(side="left", padx=0, pady=0)
        self.mode_choice = ctk.StringVar(value="")
        row = 1
        ctk.CTkLabel(
            self, text=self.tr("main_mode"), font=("Arial", 12, "italic"),
            width=360, height=18,
        ).grid(row=row, column=0, pady=(2, 7), sticky="w")
        row += 1
        self.menu_radio(self, MODES[0][0], self.tr(MODES[0][1]), row)
        row += 1
        ctk.CTkLabel(
            self, text=self.tr("additional"), font=("Arial", 12, "italic"),
            width=360, height=18,
        ).grid(row=row, column=0, pady=(9, 9), sticky="w")
        row += 1
        for key, title_key in MODES[1:]:
            if key in {"merge_pdf_mode", "images_to_pdf_mode"}:
                self.menu_separator(row)
                row += 1
            self.menu_radio(self, key, self.tr(title_key), row)
            row += 1
        about = ctk.CTkLabel(
            self, text=self.tr("about"), text_color=GREEN,
            font=("Arial", 14, "bold"), cursor="hand2", width=360, height=20,
        )
        about.grid(row=row, column=0, pady=(10, 5), sticky="w")
        about.bind("<Button-1>", lambda _event: self.show_about())
        about.bind("<Enter>", lambda _event: about.configure(text_color="#106A43"))
        about.bind("<Leave>", lambda _event: about.configure(text_color=GREEN))
        self.after_idle(lambda widget=about: self.fit_menu_window(widget))

    def fit_menu_window(self, bottom_widget=None, settle: bool = True) -> None:
        """End the main window shortly below its last visible control."""
        self.update_idletasks()
        content_bottom = (
            bottom_widget.winfo_y() + bottom_widget.winfo_height()
            if bottom_widget is not None and bottom_widget.winfo_exists()
            else self.winfo_reqheight()
        )
        scale = self._get_window_scaling()
        width = max(360, round(self.winfo_reqwidth() / scale) + 24)
        height = max(360, round(content_bottom / scale) + 16)
        height = min(height, self.winfo_screenheight() - 40)
        self.geometry(f"{width}x{height}+400+10")
        if settle:
            self.after_idle(
                lambda widget=bottom_widget: self.fit_menu_window(widget, False)
            )

    def menu_separator(self, row: int) -> None:
        width, height = 100, 6
        self._menu_separator_image = self.make_diamond_image(width, height)
        separator = ctk.CTkLabel(
            self, text="", image=self._menu_separator_image,
            width=width, height=height,
        )
        separator.grid(row=row, column=0, padx=(130, 0), pady=(8, 3), sticky="w")

    @staticmethod
    def make_diamond_image(width: int, height: int) -> ctk.CTkImage:
        """Create an antialiased flattened diamond in logical UI pixels."""
        scale = 4
        source = Image.new("RGBA", (width * scale, height * scale), (0, 0, 0, 0))
        draw = ImageDraw.Draw(source)
        draw.polygon(
            [(0, height * scale // 2), (width * scale // 2, 0),
             (width * scale, height * scale // 2),
             (width * scale // 2, height * scale)],
            fill=RED,
        )
        separator_image = source.resize(
            (width * 2, height * 2), Image.Resampling.LANCZOS
        )
        return ctk.CTkImage(
            light_image=separator_image, dark_image=separator_image,
            size=(width, height),
        )

    def menu_radio(self, parent, key: str, title: str, row: int) -> None:
        ctk.CTkRadioButton(
            parent, text=title, variable=self.mode_choice, value=key,
            height=22, radiobutton_width=17, radiobutton_height=17,
            font=("Arial", 13),
            command=lambda value=key: self.show_mode(value),
        ).grid(
            row=row, column=0, padx=45,
            pady=1, sticky="w",
        )

    def show_about(self) -> None:
        self.clear()
        self.mode = "about"
        self.geometry("620x700+400+10")
        self.back_button()
        ctk.CTkLabel(
            self, text=self.tr("about"), font=("Arial", 18, "bold"),
        ).grid(row=0, column=0, pady=(22, 0))
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, padx=70, pady=10, sticky="n")
        try:
            self._about_icon = ctk.CTkImage(
                light_image=Image.open(resource_path("label.png")),
                dark_image=Image.open(resource_path("label.png")), size=(92, 92),
            )
            ctk.CTkLabel(content, text="", image=self._about_icon).pack(pady=(4, 0))
        except OSError:
            pass
        ctk.CTkLabel(content, text="ScanItEasy", font=("Arial", 25, "bold")).pack(pady=(10, 0))
        ctk.CTkLabel(
            content, text=f'{self.tr("version")} {APP_VERSION}', text_color="gray"
        ).pack(pady=(2, 0))
        ctk.CTkLabel(content, text=self.tr("developed_by"), text_color="gray").pack(pady=(10, 0))
        ctk.CTkLabel(content, text=self.tr("about_intro"), wraplength=460).pack(pady=(12, 6))
        ctk.CTkLabel(
            content, text=self.tr("contacts"), font=("Arial", 17, "bold"),
        ).pack(anchor="w", pady=(18, 8))
        self.about_card(content, "✉", "Email", "contact@formacode.dev", EMAIL_URL)
        self.about_card(content, "➤", "Telegram", "@PavelCraftRobot", TELEGRAM_URL)
        self.about_card(
            content, "?", self.tr("instruction"), self.tr("instruction_hint"),
            GUIDE_URLS[self.language],
        )

    def about_card(self, parent, icon: str, title: str, detail: str, url: str) -> None:
        card = ctk.CTkFrame(parent, width=470, height=66, corner_radius=12)
        card.pack(fill="x", pady=5)
        card.pack_propagate(False)
        ctk.CTkLabel(card, text=icon, width=42, text_color=GREEN,
                     font=("Arial", 23, "bold")).pack(side="left", padx=(12, 4))
        copy = ctk.CTkFrame(card, fg_color="transparent")
        copy.pack(side="left", fill="both", expand=True, pady=8)
        ctk.CTkLabel(copy, text=title, font=("Arial", 14, "bold")).pack(anchor="w")
        ctk.CTkLabel(copy, text=detail, text_color="gray", font=("Arial", 12)).pack(anchor="w")
        ctk.CTkLabel(card, text="›", width=30, text_color="gray",
                     font=("Arial", 28)).pack(side="right", padx=8)
        self.bind_link_tree(card, url)

    def bind_link_tree(self, widget, url: str) -> None:
        widget.bind("<Button-1>", lambda _event, target=url: webbrowser.open(target))
        widget.configure(cursor="hand2")
        for child in widget.winfo_children():
            self.bind_link_tree(child, url)

    def show_mode(self, mode: str) -> None:
        self.clear()
        self.mode = mode
        self.minsize(360, 240)
        self.grid_rowconfigure(2, weight=0)
        self.selected_images = []
        self.merge_files = []
        self.merge_entries = {}
        self.image_row_widgets = {}
        self.merge_row_widgets = {}
        self._file_drag = None
        self._drag_indicator = None
        self._drag_insertion = None
        title = self.tr(dict(MODES)[mode])
        self.back_button()
        ctk.CTkLabel(self, text=title, font=("Arial", 17, "bold")).grid(
            row=1, column=0, padx=70, pady=5, sticky="w"
        )
        self.grid_rowconfigure(2, weight=1)
        self.form = AutoHideScrollableFrame(
            self, fg_color="transparent", width=500, height=40,
        )
        self.form.grid(row=2, column=0, padx=(55, 15), pady=0, sticky="nsew")
        self.form_row = 0
        if mode in {"custom_mode", "convert_to_pdf_mode",
                    "convert_docx_to_png_mode", "convert_docx_to_jpg_mode"}:
            self.file_picker("docx_path", self.tr("select_docx"), self.file_types("docx"))
            self.page_entry("pages", self.tr("select_pages"))
            if mode == "custom_mode":
                self.radio_group("color", self.tr("ribbon_color"), ["red", "blue"], "red",
                                 labels={"red": self.tr("red"), "blue": self.tr("blue")})
            if mode == "custom_mode":
                self.file_picker("last_page_path", self.tr("select_last_part"),
                                 [(self.tr("pdf_or_image_type"), "*.pdf *.png *.jpg *.jpeg")], optional=True)
                self.page_entry("last_page_pages", self.tr("last_part_pages"), optional=True)
                self.set_field_visible("last_page_pages", False)
            elif mode in {"convert_docx_to_png_mode", "convert_docx_to_jpg_mode"}:
                self.number_entry("dpi", self.tr("resolution_dpi"), "200")
                if mode == "convert_docx_to_jpg_mode":
                    self.number_entry("jpeg_quality", self.tr("jpg_quality"), "85")
        elif mode in {"grayscale_mode", "convert_to_png_mode", "convert_to_jpg_mode"}:
            self.file_picker("pdf_path", self.tr("select_pdf"), self.file_types("pdf"))
            self.page_entry("pages", self.tr("select_pages"))
            if mode in {"convert_to_png_mode", "convert_to_jpg_mode"}:
                self.number_entry("dpi", self.tr("resolution_dpi"), "200")
            if mode == "convert_to_jpg_mode":
                self.number_entry("jpeg_quality", self.tr("jpg_quality"), "85")
        elif mode == "compress_mode":
            self.merge_picker(self.tr("select_pdfs_to_compress"))
            self.compression_output_controls()
            self.radio_group("profile", self.tr("quality"), ["high", "balanced", "small"],
                             "balanced", labels={"high": self.tr("quality_high"),
                             "balanced": self.tr("quality_balanced"),
                             "small": self.tr("quality_small")})
        elif mode == "merge_pdf_mode":
            self.merge_picker()
        elif mode in {"images_to_pdf_mode", "compress_images_mode", "resize_images_mode"}:
            self.multi_image_picker()
            if mode == "compress_images_mode":
                self.checkbox("lossy", self.tr("lossy_images"), False,
                              command=self.toggle_image_compression)
                self.number_entry("jpeg_quality", self.tr("jpg_quality"), "85")
                self.radio_group("png_colors", self.tr("png_colors"),
                                 ["256", "128", "64"], "256")
                self.set_field_visible("jpeg_quality", False)
                self.set_field_visible("png_colors", False)
            elif mode == "resize_images_mode":
                self.number_entry("resize_percent", self.tr("resize_percent"), "100")
                self.number_entry("jpeg_quality", self.tr("jpg_save_quality"), "85")
                self.set_field_visible("jpeg_quality", False)
        if self.produces_pdf(mode) and mode != "compress_mode":
            self.checkbox("is_compression_needed", self.tr("compress_pdf"), False)
        if self.produces_pdf(mode):
            self.acrobat_controls(self.form)
        self.status_row = self.form_row
        self.status_frame = ctk.CTkFrame(self.form, fg_color="transparent")
        status_background = self._apply_appearance_mode(self.cget("fg_color"))
        self.status = tk.Text(
            self.status_frame, width=52, height=1, wrap="word",
            borderwidth=0, highlightthickness=0, relief="flat",
            bg=status_background, fg=self._apply_appearance_mode(TEXT),
            insertbackground=status_background, font=("Arial", 14),
            cursor="arrow", takefocus=False,
        )
        self.status.pack(side="left")
        self.status.bind("<MouseWheel>", lambda _event: "break")
        self.status.bind("<Button-4>", lambda _event: "break")
        self.status.bind("<Button-5>", lambda _event: "break")
        self.status.tag_configure("ghostscript", foreground=RED, underline=True)
        self.status.tag_bind(
            "ghostscript", "<Enter>",
            lambda _event: (
                self.status.configure(cursor="hand2"),
                self.status.tag_configure("ghostscript", foreground=DARK_RED),
            ),
        )
        self.status.tag_bind(
            "ghostscript", "<Leave>",
            lambda _event: (
                self.status.configure(cursor="arrow"),
                self.status.tag_configure("ghostscript", foreground=RED),
            ),
        )
        self.status.tag_bind(
            "ghostscript", "<Button-1>",
            lambda _event: webbrowser.open(GHOSTSCRIPT_GUIDE_URLS[self.language])
        )
        self.form_row += 1
        self.progress_row = self.form_row
        self.progress = ctk.CTkProgressBar(self.form, width=300, height=25)
        self.progress.set(0)
        self.form_row += 1
        action = {"custom_mode": self.tr("scan"),
                  "compress_mode": self.tr("compress"), "grayscale_mode": self.tr("change"),
                  "resize_images_mode": self.tr("change"),
                  "merge_pdf_mode": self.tr("merge")}.get(mode, self.tr("convert"))
        self.action_frame = ctk.CTkFrame(self.form, fg_color="transparent", width=300)
        self.action_frame.grid(row=self.form_row, column=0, pady=5, sticky="w")
        self.run_button = ctk.CTkButton(
            self.action_frame, text=action, width=140, fg_color=RED,
            hover_color=DARK_RED, command=self.start_job,
        )
        self.run_button.grid(row=0, column=0, sticky="w")
        self.open_folder_button = ctk.CTkButton(
            self.action_frame, text=self.tr("open_folder"), width=140,
            command=self.open_output_folder,
        )
        self.form_row += 1
        self.continue_row = self.form_row
        self.choose_acrobat_row = self.form_row + 1
        self.form_row += 3
        self.continue_button = ctk.CTkButton(self.form, text=self.tr("continue"),
                                             command=self.retry_job)
        self.choose_acrobat_button = ctk.CTkButton(self.form, text=self.tr("specify_path"),
                                                    command=self.choose_acrobat_folder)
        self.after_idle(self.fit_window_to_content)

    def back_button(self) -> None:
        light = Image.open(resource_path("light.png"))
        dark = Image.open(resource_path("dark.png"))
        self._back_light = ctk.CTkImage(
            light_image=light, dark_image=light, size=(24, 29),
        )
        self._back_dark = ctk.CTkImage(
            light_image=dark, dark_image=dark, size=(24, 29),
        )
        background = self._apply_appearance_mode(self.cget("fg_color"))
        button = ctk.CTkButton(
            self, text="", image=self._back_light,
            width=31, height=36, border_spacing=0,
            fg_color="transparent", hover_color=background,
            cursor="hand2", command=self.show_menu,
        )
        button.grid(row=0, column=0, padx=35, pady=(12, 5), sticky="w")
        button.bind("<Enter>", lambda _event: button.configure(image=self._back_dark))
        button.bind("<Leave>", lambda _event: button.configure(image=self._back_light))

    def labeled_frame(self, label: str):
        frame = ctk.CTkFrame(self.form, fg_color="transparent")
        frame.grid(row=self.form_row, column=0, pady=5, sticky="w")
        self.form_row += 1
        ctk.CTkLabel(frame, text=label).grid(row=0, column=0, sticky="w")
        return frame

    def file_picker(self, key: str, title: str, types, optional: bool = False) -> None:
        frame = self.labeled_frame(title)
        value = ctk.StringVar()
        label = ctk.CTkLabel(frame, textvariable=value, wraplength=470, text_color=GREEN)
        button = ctk.CTkButton(
            frame, text=self.tr("choose_file"),
            command=lambda: self.select_file(key, value, types, label),
        )
        button.grid(row=1, column=0, sticky="w")
        self.fields[key] = value

    def select_file(self, key: str, variable: ctk.StringVar, types,
                    label: ctk.CTkLabel) -> None:
        filename = filedialog.askopenfilename(filetypes=types)
        if filename:
            variable.set(filename)
            label.grid(row=2, column=0, pady=0, sticky="w")
            if key == "last_page_path":
                self.set_field_visible("last_page_pages", False)
                show_pages = is_multipage_pdf(filename)
                if not show_pages:
                    pages = self.fields["last_page_pages"]
                    pages.delete(0, "end")
                self.set_field_visible("last_page_pages", show_pages)
            self.after_idle(self.fit_window_to_content)

    def page_entry(self, key: str, title: str, optional: bool = False) -> None:
        frame = ctk.CTkFrame(self.form, fg_color="transparent")
        frame.grid(row=self.form_row, column=0, pady=0, sticky="w")
        self.form_row += 1
        ctk.CTkLabel(frame, text=title).grid(row=0, column=0, pady=0, sticky="w")
        ctk.CTkLabel(
            frame, text=self.tr("all_pages_hint"),
            text_color=GREEN, font=("Arial", 12, "italic"), justify="center",
        ).grid(row=1, column=0, pady=0, sticky="w")
        entry = ctk.CTkEntry(
            frame, width=470, fg_color="lightgray", text_color="black",
            placeholder_text_color="gray", corner_radius=10,
            font=("Arial", 12),
            placeholder_text=self.tr("pages_placeholder"),
        )
        entry.grid(row=2, column=0, pady=5, sticky="w")
        self.fields[key] = entry
        self.field_frames[key] = frame

    def number_entry(self, key: str, title: str, default: str) -> None:
        frame = self.labeled_frame(title)
        entry = ctk.CTkEntry(frame, width=180)
        entry.insert(0, default)
        entry.grid(row=1, column=0, pady=5, sticky="w")
        self.fields[key] = entry
        self.field_frames[key] = frame

    def checkbox(self, key: str, text: str, default: bool, command=None) -> None:
        variable = ctk.BooleanVar(value=default)
        checkbox = ctk.CTkCheckBox(self.form, text=text, variable=variable, command=command)
        checkbox.grid(row=self.form_row, column=0, pady=5, sticky="w")
        self.form_row += 1
        self.fields[key] = variable

    def radio_group(self, key: str, title: str, values: list[str], default: str,
                    labels: dict[str, str] | None = None) -> None:
        frame = self.labeled_frame(title)
        mapping = labels or {value: value for value in values}
        variable = ctk.StringVar(value=default)
        for row, value in enumerate(values, 1):
            ctk.CTkRadioButton(
                frame, text=mapping[value], variable=variable, value=value,
            ).grid(row=row, column=0, pady=4, sticky="w")
        self.fields[key] = variable
        self.field_frames[key] = frame

    def set_field_state(self, key: str, enabled: bool) -> None:
        item = self.fields.get(key)
        if hasattr(item, "configure"):
            item.configure(state="normal" if enabled else "disabled")
        frame = self.field_frames.get(key)
        if frame:
            for widget in frame.winfo_children():
                if isinstance(widget, ctk.CTkRadioButton):
                    widget.configure(state="normal" if enabled else "disabled")

    def set_field_visible(self, key: str, visible: bool) -> None:
        frame = self.field_frames.get(key)
        if frame is None:
            return
        if visible:
            frame.grid()
        else:
            frame.grid_remove()
        self.after_idle(self.fit_window_to_content)

    def toggle_image_compression(self) -> None:
        enabled = bool(self.fields["lossy"].get())
        has_jpeg = any(Path(name).suffix.lower() in {".jpg", ".jpeg"} for name in self.selected_images)
        has_png = any(Path(name).suffix.lower() == ".png" for name in self.selected_images)
        self.set_field_visible("jpeg_quality", has_jpeg)
        self.set_field_visible("png_colors", has_png)
        self.set_field_state("jpeg_quality", enabled and has_jpeg)
        self.set_field_state("png_colors", enabled and has_png)

    def multi_image_picker(self) -> None:
        frame = self.labeled_frame(self.tr("select_images"))
        ctk.CTkButton(frame, text=self.tr("choose_images"), command=self.select_images).grid(
            row=1, column=0, sticky="w"
        )
        self.image_rows = ctk.CTkFrame(frame, fg_color="transparent")

    def merge_picker(self, title: str | None = None) -> None:
        frame = self.labeled_frame(title or self.tr("select_pdfs"))
        ctk.CTkButton(frame, text=self.tr("add_files"), command=self.select_merge_files).grid(
            row=1, column=0, sticky="w"
        )
        self.merge_rows = ctk.CTkFrame(frame, fg_color="transparent")

    def select_merge_files(self) -> None:
        for filename in filedialog.askopenfilenames(filetypes=self.file_types("pdf")):
            if filename not in self.merge_files:
                self.merge_files.append(filename)
        self.render_merge_rows()
        self.update_compression_output_controls()

    def compression_output_controls(self) -> None:
        self.compression_folder_frame = ctk.CTkFrame(
            self.form, fg_color="transparent",
        )
        self.compression_folder_frame.grid(
            row=self.form_row, column=0, pady=5, sticky="w",
        )
        self.form_row += 1
        ctk.CTkLabel(
            self.compression_folder_frame,
            text=self.tr("compression_different_folders"),
            wraplength=470, justify="left",
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            self.compression_folder_frame,
            text=self.tr("choose_output_folder"),
            command=self.select_compression_output_folder,
        ).grid(row=1, column=0, pady=(5, 0), sticky="w")
        output_folder = ctk.StringVar(value="")
        self.compression_folder_selection = ctk.CTkFrame(
            self.compression_folder_frame, fg_color="transparent", width=470,
        )
        self.compression_folder_selection.grid(
            row=2, column=0, pady=(4, 0), sticky="ew",
        )
        self.compression_folder_selection.grid_columnconfigure(0, weight=1)
        self.compression_folder_label = ctk.CTkLabel(
            self.compression_folder_selection, textvariable=output_folder,
            text_color=GREEN, wraplength=400, justify="left", anchor="w",
        )
        self.compression_folder_label.grid(row=0, column=0, sticky="ew")
        ctk.CTkButton(
            self.compression_folder_selection, text="×", width=28, height=26,
            fg_color="transparent", text_color="red", hover_color=DARK_RED,
            command=self.clear_compression_output_folder,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ne")
        self.compression_folder_selection.grid_remove()
        self.fields["compression_output_folder"] = output_folder
        self.compression_folder_frame.grid_remove()

    def select_compression_output_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder:
            self.fields["compression_output_folder"].set(folder)
            self.compression_folder_selection.grid()
            self.after_idle(self.fit_window_to_content)

    def clear_compression_output_folder(self) -> None:
        self.fields["compression_output_folder"].set("")
        self.compression_folder_selection.grid_remove()
        self.after_idle(self.fit_window_to_content)

    def update_compression_output_controls(self) -> None:
        if self.mode != "compress_mode" or not hasattr(
            self, "compression_folder_frame",
        ):
            return
        folders = {
            os.path.normcase(str(Path(filename).resolve().parent))
            for filename in self.merge_files
        }
        if len(folders) > 1:
            self.compression_folder_frame.grid()
        else:
            self.compression_folder_frame.grid_remove()
            self.clear_compression_output_folder()
        self.after_idle(self.fit_window_to_content)

    def render_merge_rows(self) -> None:
        saved = {name: entry.get() for name, entry in self.merge_entries.items()}
        for child in self.merge_rows.winfo_children():
            child.destroy()
        self.merge_entries.clear()
        self.merge_row_widgets.clear()
        if not self.merge_files:
            self.merge_rows.grid_forget()
            return
        self.merge_rows.grid(row=2, column=0, pady=(6, 0), sticky="w")
        for index, filename in enumerate(self.merge_files):
            row = ctk.CTkFrame(self.merge_rows, fg_color="transparent", width=470)
            row.grid(row=index * 2, column=0, sticky="ew", pady=(2, 0))
            draggable = self.mode == "merge_pdf_mode"
            name_column = 1 if draggable else 0
            button_column = 2 if draggable else 1
            name_width = 370 if draggable else 400
            row.grid_columnconfigure(name_column, weight=1)
            if draggable:
                handle = ctk.CTkLabel(
                    row, text="⠿", width=24, cursor="fleur", text_color="gray",
                    font=("Segoe UI Symbol", 18),
                )
                handle.grid(row=0, column=0, padx=(0, 6), sticky="n")
                self.bind_file_drag(handle, "merge", filename, row)
            ctk.CTkLabel(
                row, text=Path(filename).name, width=name_width, anchor="w",
                justify="left", wraplength=name_width,
            ).grid(row=0, column=name_column, sticky="ew")
            ctk.CTkButton(
                row, text="×", width=28, height=26, fg_color="transparent",
                text_color="red", hover_color=DARK_RED,
                command=lambda item=filename: self.remove_merge_file(item),
            ).grid(row=0, column=button_column, padx=(8, 0), sticky="ne")
            self.merge_row_widgets[filename] = row
            if self.mode == "merge_pdf_mode":
                entry = ctk.CTkEntry(
                    self.merge_rows, width=470,
                    placeholder_text=self.tr("merge_pages_placeholder"),
                )
                entry.grid(row=index * 2 + 1, column=0, sticky="w", pady=(0, 5))
                if saved.get(filename):
                    entry.insert(0, saved[filename])
                self.merge_entries[filename] = entry
        self.after_idle(self.fit_window_to_content)

    def remove_merge_file(self, filename: str) -> None:
        if filename in self.merge_files:
            self.merge_files.remove(filename)
        self.render_merge_rows()
        self.update_compression_output_controls()

    def select_images(self) -> None:
        files = list(filedialog.askopenfilenames(filetypes=self.file_types("images")))
        if files:
            for filename in files:
                if filename not in self.selected_images:
                    self.selected_images.append(filename)
            self.render_image_rows()

    def render_image_rows(self) -> None:
        for child in self.image_rows.winfo_children():
            child.destroy()
        self.image_row_widgets.clear()
        if not self.selected_images:
            self.image_rows.grid_forget()
        else:
            self.image_rows.grid(row=2, column=0, pady=(6, 0), sticky="w")
            self.image_rows.grid_columnconfigure(0, weight=1)
            for index, filename in enumerate(self.selected_images):
                row = ctk.CTkFrame(
                    self.image_rows, fg_color="transparent", width=470,
                )
                row.grid(row=index, column=0, sticky="ew", pady=(1, 1))
                draggable = self.mode == "images_to_pdf_mode"
                name_column = 1 if draggable else 0
                button_column = 2 if draggable else 1
                name_width = 370 if draggable else 400
                row.grid_columnconfigure(name_column, weight=1)
                if draggable:
                    handle = ctk.CTkLabel(
                        row, text="⠿", width=24, cursor="fleur",
                        text_color="gray", font=("Segoe UI Symbol", 18),
                    )
                    handle.grid(row=0, column=0, padx=(0, 6), sticky="n")
                    self.bind_file_drag(handle, "image", filename, row)
                ctk.CTkLabel(
                    row, text=Path(filename).name, width=name_width, anchor="w",
                    justify="left", wraplength=name_width,
                ).grid(row=0, column=name_column, sticky="ew")
                ctk.CTkButton(
                    row, text="×", width=28, height=26, fg_color="transparent",
                    text_color="red", hover_color=DARK_RED,
                    command=lambda item=filename: self.remove_image(item),
                ).grid(row=0, column=button_column, padx=(8, 0), sticky="ne")
                self.image_row_widgets[filename] = row
        if self.mode == "compress_images_mode":
            self.toggle_image_compression()
        elif self.mode == "resize_images_mode":
            has_jpeg = any(
                Path(name).suffix.lower() in {".jpg", ".jpeg"}
                for name in self.selected_images
            )
            self.set_field_visible("jpeg_quality", has_jpeg)
        self.after_idle(self.fit_window_to_content)

    def remove_image(self, filename: str) -> None:
        if filename in self.selected_images:
            self.selected_images.remove(filename)
        self.render_image_rows()

    def bind_file_drag(self, handle, kind: str, filename: str, row) -> None:
        handle.bind(
            "<ButtonPress-1>",
            lambda event: self.start_file_drag(event, kind, filename, row),
        )
        handle.bind("<B1-Motion>", self.move_file_drag)
        handle.bind("<ButtonRelease-1>", self.finish_file_drag)

    def start_file_drag(self, event, kind: str, filename: str, row) -> str:
        self._file_drag = (kind, filename, row)
        row.configure(fg_color=("gray85", "gray25"))
        self.update_file_drag_indicator(event.y_root)
        return "break"

    def move_file_drag(self, event) -> str:
        if self._file_drag is None:
            return "break"
        canvas = self.form._parent_canvas
        pointer_y = event.y_root
        top = canvas.winfo_rooty()
        bottom = top + canvas.winfo_height()
        if pointer_y < top + 30:
            canvas.yview_scroll(-1, "units")
        elif pointer_y > bottom - 30:
            canvas.yview_scroll(1, "units")
        self.update_idletasks()
        self.update_file_drag_indicator(pointer_y)
        return "break"

    def update_file_drag_indicator(self, pointer_y: int) -> None:
        if self._file_drag is None:
            return
        kind, _filename, _row = self._file_drag
        items = self.merge_files if kind == "merge" else self.selected_images
        rows = self.merge_row_widgets if kind == "merge" else self.image_row_widgets
        available = [(item, rows[item]) for item in items if item in rows]
        if not available:
            return
        current_insertion = sum(
            widget.winfo_rooty() + widget.winfo_height() / 2 <= pointer_y
            for _item, widget in available
        )
        if (
            current_insertion == self._drag_insertion
            and self._drag_indicator is not None
            and self._drag_indicator.winfo_exists()
        ):
            return
        self.reset_file_drag_spacing(kind)
        self.update_idletasks()
        insertion = sum(
            widget.winfo_rooty() + widget.winfo_height() / 2 <= pointer_y
            for _item, widget in available
        )
        parent = self.merge_rows if kind == "merge" else self.image_rows
        gap = 20
        if insertion < len(available):
            available[insertion][1].grid_configure(pady=(gap, 0))
        elif kind == "merge" and available[-1][0] in self.merge_entries:
            last_entry = self.merge_entries[available[-1][0]]
            last_entry.grid_configure(pady=(0, gap))
        else:
            last_row = available[-1][1]
            last_row.grid_configure(pady=(1, gap))
        self.update_idletasks()
        widget_scale = available[0][1]._get_widget_scaling()
        if insertion < len(available):
            line_y = available[insertion][1].winfo_y() / widget_scale - gap / 2
        elif kind == "merge" and available[-1][0] in self.merge_entries:
            last_entry = self.merge_entries[available[-1][0]]
            line_y = (
                (last_entry.winfo_y() + last_entry.winfo_height()) / widget_scale
                + gap / 2
            )
        else:
            last_row = available[-1][1]
            line_y = (
                (last_row.winfo_y() + last_row.winfo_height()) / widget_scale
                + gap / 2
            )
        if self._drag_indicator is None or not self._drag_indicator.winfo_exists():
            self._drag_indicator_image = self.make_diamond_image(100, 6)
            self._drag_indicator = ctk.CTkLabel(
                parent, text="", image=self._drag_indicator_image,
                width=100, height=6,
            )
        self._drag_indicator.place(
            x=(470 - 100) / 2,
            y=line_y - 3,
        )
        self._drag_indicator.lift()
        self._drag_insertion = insertion

    def reset_file_drag_spacing(self, kind: str) -> None:
        if kind == "merge":
            for row in self.merge_row_widgets.values():
                if row.winfo_exists():
                    row.grid_configure(pady=(2, 0))
            for entry in self.merge_entries.values():
                if entry.winfo_exists():
                    entry.grid_configure(pady=(0, 5))
        else:
            for row in self.image_row_widgets.values():
                if row.winfo_exists():
                    row.grid_configure(pady=(1, 1))

    def clear_file_drag_indicator(self, kind: str) -> None:
        self.reset_file_drag_spacing(kind)
        if self._drag_indicator is not None and self._drag_indicator.winfo_exists():
            self._drag_indicator.destroy()
        self._drag_indicator = None

    def finish_file_drag(self, event) -> str:
        if self._file_drag is None:
            return "break"
        kind, filename, dragged_row = self._file_drag
        self.update_file_drag_indicator(event.y_root)
        insertion = self._drag_insertion
        self._file_drag = None
        self._drag_insertion = None
        self.clear_file_drag_indicator(kind)
        items = self.merge_files if kind == "merge" else self.selected_images
        if filename not in items or insertion is None:
            if dragged_row.winfo_exists():
                dragged_row.configure(fg_color="transparent")
            return "break"
        source = items.index(filename)
        item = items.pop(source)
        if insertion > source:
            insertion -= 1
        insertion = min(insertion, len(items))
        items.insert(insertion, item)
        if source != insertion:
            if kind == "merge":
                self.render_merge_rows()
            else:
                self.render_image_rows()
        moved_row = (
            self.merge_row_widgets.get(filename)
            if kind == "merge" else self.image_row_widgets.get(filename)
        )
        if moved_row is not None and moved_row.winfo_exists():
            moved_row.configure(fg_color=("gray85", "gray25"))
            self.after(500, lambda row=moved_row: self.clear_drop_highlight(row))
        return "break"

    @staticmethod
    def clear_drop_highlight(row) -> None:
        if row.winfo_exists():
            row.configure(fg_color="transparent")

    def fit_window_to_content(self) -> None:
        if not hasattr(self, "form") or not self.form.winfo_exists():
            return
        self.update_idletasks()
        content = self.form._parent_canvas.bbox("all")
        content_height = 0 if content is None else content[3] - content[1]
        scale = self._get_window_scaling()
        # Keep only a small margin below the final visible control.
        # Include the title area and leave a modest gap below the last control.
        # The extra reserve also prevents one-pixel rounding differences at
        # 100% Windows scaling from causing an unnecessary scrollbar.
        chrome_height = 120
        desired_height = max(280, round(content_height / scale) + chrome_height)
        # Tk already reports the screen height in logical pixels here.
        screen_limit = max(360, self.winfo_screenheight() - 70)
        height = min(desired_height, screen_limit)
        self.geometry(f"650x{height}+400+10")
        self.after_idle(self.form.update_scrollbar_visibility)

    def acrobat_controls(self, parent) -> None:
        saved_open = bool(self.preferences.get("open_result_file", True))
        saved_method = (
            "acrobat" if self.preferences.get("open_in_acrobat", False)
            else "default"
        )
        saved_page = str(self.preferences.get("acrobat_page", 1))
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=self.form_row, column=0, pady=5, sticky="w")
        self.form_row += 1

        enabled = ctk.BooleanVar(value=saved_open)
        details = ctk.CTkFrame(frame, fg_color="transparent")
        ctk.CTkCheckBox(
            frame, text=self.tr("open_result"), variable=enabled,
            command=lambda: self.open_result_toggled(enabled, details),
        ).grid(row=0, column=0, sticky="w")

        method = ctk.StringVar(value=saved_method)
        ctk.CTkRadioButton(
            details, text=self.tr("open_default"), variable=method,
            value="default", command=lambda: self.open_method_changed(method, page),
        ).grid(row=0, column=0, pady=(7, 4), sticky="w")
        acrobat_row = ctk.CTkFrame(details, fg_color="transparent")
        acrobat_row.grid(row=1, column=0, pady=4, sticky="w")
        ctk.CTkRadioButton(
            acrobat_row, text=self.tr("open_acrobat"), variable=method,
            value="acrobat", command=lambda: self.open_method_changed(method, page),
        ).pack(side="left")
        page = ctk.CTkEntry(acrobat_row, width=52)
        page.insert(0, saved_page)
        page.pack(side="left", padx=(6, 0))
        page.configure(state="normal" if saved_method == "acrobat" else "disabled")
        if saved_open:
            details.grid(row=1, column=0, padx=(24, 0), sticky="w")

        self.fields["open_result_file"] = enabled
        self.fields["open_method"] = method
        self.fields["acrobat_page"] = page

    def open_result_toggled(self, enabled, details) -> None:
        if enabled.get():
            details.grid(row=1, column=0, padx=(24, 0), sticky="w")
        else:
            details.grid_remove()
        self.after_idle(self.fit_window_to_content)

    def open_method_changed(self, method, page) -> None:
        if method.get() == "acrobat":
            page.configure(state="normal")
            page.delete(0, "end")
            page.insert(0, "1")
        else:
            page.configure(state="disabled")

    @staticmethod
    def produces_pdf(mode: str) -> bool:
        return mode in {"custom_mode", "convert_to_pdf_mode",
                        "grayscale_mode", "compress_mode", "images_to_pdf_mode",
                        "merge_pdf_mode"}

    def field_value(self, key: str, default=None):
        item = self.fields.get(key)
        if item is None:
            return default
        return item.get()

    def collect_job(self) -> dict | None:
        settings: dict = {"mode": self.mode}
        if self.mode in {"custom_mode", "convert_to_pdf_mode",
                         "convert_docx_to_png_mode", "convert_docx_to_jpg_mode"}:
            path = self.field_value("docx_path", "")
            if not path:
                self.show_status(self.tr("choose_docx_error"), error=True); return None
            settings.update(docx_path=path, pages=self.field_value("pages", ""),
                            last_page_path=self.field_value("last_page_path", ""),
                            last_page_pages=self.field_value("last_page_pages", ""),
                            color=self.field_value("color", "red"),
                            dpi=self.field_value("dpi", "200"),
                            jpeg_quality=self.field_value("jpeg_quality", "85"),
                            is_compression_needed=bool(self.field_value("is_compression_needed", False)))
        elif self.mode == "merge_pdf_mode":
            if len(self.merge_files) < 2:
                self.show_status(self.tr("choose_two_pdfs_error"), error=True); return None
            pages = [self.merge_entries[name].get() for name in self.merge_files]
            try:
                for value in pages:
                    if value:
                        parse_page_sequence(value)
            except ValueError:
                self.show_status(self.tr("check_pages"), error=True); return None
            output = filedialog.asksaveasfilename(
                defaultextension=".pdf", filetypes=self.file_types("pdf"),
                initialfile="merged_output.pdf",
            )
            if not output:
                return None
            settings = {
                "mode": "merge_pdf",
                "parameters": (list(self.merge_files), pages, output),
                "is_compression_needed": bool(
                    self.field_value("is_compression_needed", False)
                ),
            }
        elif self.mode == "compress_mode":
            if not self.merge_files:
                self.show_status(self.tr("choose_pdf_error"), error=True); return None
            settings.update(
                pdf_paths=list(self.merge_files),
                profile=self.field_value("profile", "balanced"),
                output_folder=self.field_value("compression_output_folder", ""),
            )
        elif self.mode in {"grayscale_mode", "convert_to_png_mode", "convert_to_jpg_mode"}:
            path = self.field_value("pdf_path", "")
            if not path:
                self.show_status(self.tr("choose_pdf_error"), error=True); return None
            settings.update(pdf_path=path, pages=self.field_value("pages", ""),
                            profile=self.field_value("profile", "balanced"),
                            dpi=self.field_value("dpi", "200"),
                            jpeg_quality=self.field_value("jpeg_quality", "85"),
                            is_compression_needed=bool(
                                self.field_value("is_compression_needed", False)
                            ))
        else:
            if not self.selected_images:
                self.show_status(self.tr("choose_image_error"), error=True); return None
            settings.update(image_paths=list(self.selected_images),
                            lossy=bool(self.field_value("lossy", False)),
                            jpeg_quality=self.field_value("jpeg_quality", "85"),
                            png_colors=self.field_value("png_colors", "256"),
                            resize_percent=self.field_value("resize_percent", "100"),
                            is_compression_needed=bool(
                                self.field_value("is_compression_needed", False)
                            ))
        try:
            for key in ("pages", "last_page_pages"):
                if settings.get(key): parse_page_sequence(settings[key])
            for key, low, high in (("dpi", 50, 1200), ("jpeg_quality", 1, 95),
                                   ("resize_percent", 1, 500)):
                if key in settings and settings[key] not in (None, ""):
                    value = int(settings[key])
                    if not low <= value <= high: raise ValueError
                    settings[key] = value
        except (ValueError, TypeError):
            self.show_status(self.tr("check_parameters"), error=True)
            return None
        return settings

    def start_job(self) -> None:
        job = self.collect_job()
        if job is not None:
            self.last_job = job
            self.launch(job)

    def retry_job(self) -> None:
        self.continue_button.grid_forget()
        if self.last_job:
            self.launch(self.last_job)

    def launch(self, job: dict) -> None:
        if self.worker and self.worker.is_alive(): return
        self.result = None
        self.progress_value = 0
        self.progress.set(0)
        self.progress.grid(row=self.progress_row, column=0, pady=5, sticky="w")
        self.after_idle(self.fit_window_to_content)
        self.show_status(self.tr("processing"))
        self.run_button.configure(state="disabled")
        self.open_folder_button.grid_forget(); self.choose_acrobat_button.grid_forget()
        self.continue_button.grid_forget()
        def run():
            self.result = process_job(job, lambda value: setattr(self, "progress_value", value))
        self.worker = threading.Thread(target=run, daemon=True)
        self.worker.start()
        self.after(150, self.poll_job)

    def poll_job(self) -> None:
        self.progress.set(self.progress_value / 100)
        if self.worker and self.worker.is_alive():
            self.after(150, self.poll_job); return
        self.run_button.configure(state="normal")
        result = self.result
        if result is None:
            self.show_status(self.tr("no_worker_result"), error=True); return
        if not result.success:
            if result.ghostscript_missing:
                self.show_ghostscript_status(
                    self.tr("compression_no_gain") + " " + self.tr("install_ghostscript")
                )
            else:
                self.show_status(
                    self.tr(result.message, **result.message_params), error=True
                )
            if result.error_code in {"input_open", "output_locked"}:
                self.continue_button.grid(
                    row=self.continue_row, column=0, pady=5, sticky="w"
                )
            return
        self.progress.set(1)
        text = self.tr(result.message, **result.message_params)
        if result.notices:
            text += " " + " ".join(
                self.tr(notice.code, **notice.params) for notice in result.notices
            )
        if result.ghostscript_missing:
            self.show_ghostscript_status(
                text + " " + self.tr("stronger_compression")
            )
        else:
            self.show_status(text)
        result_folders = {
            os.path.normcase(str(Path(output).resolve().parent))
            for output in result.outputs
        }
        if not (self.mode == "compress_mode" and len(result_folders) > 1):
            self.open_folder_button.grid(
                row=0, column=1, padx=(20, 0), sticky="e"
            )
        self.save_acrobat_preferences()
        if self.produces_pdf(self.mode) and result.outputs:
            self.handle_opening(result.outputs[0])
            if (
                self.mode == "compress_mode"
                and self.field_value("open_result_file", True)
            ):
                prefer = self.field_value("open_method", "default") == "acrobat"
                page = self.field_value("acrobat_page", "1")
                for output in result.outputs[1:]:
                    open_result(
                        output, prefer, page,
                        self.preferences.get("acrobat_path"),
                    )

    def save_acrobat_preferences(self) -> None:
        if "open_result_file" not in self.fields: return
        self.preferences["open_result_file"] = bool(
            self.field_value("open_result_file", True)
        )
        self.preferences["open_in_acrobat"] = (
            self.field_value("open_method", "default") == "acrobat"
        )
        page = self.field_value("acrobat_page", "1") or "1"
        try: page = int(page)
        except ValueError: page = 1
        self.preferences["acrobat_page"] = page
        save_settings(self.preferences)

    def handle_opening(self, output: str) -> None:
        if not self.field_value("open_result_file", True):
            return
        prefer = self.field_value("open_method", "default") == "acrobat"
        page = self.field_value("acrobat_page", "1")
        opened = open_result(output, prefer, page, self.preferences.get("acrobat_path"))
        if prefer and opened.acrobat_missing:
            self.show_status(self.tr("acrobat_missing") + (self.tr("opened_default") if opened.opened else ""), error=True)
            self.choose_acrobat_button.grid(
                row=self.choose_acrobat_row, column=0, pady=5, sticky="w"
            )
        elif not prefer and opened.default_app_missing:
            self.show_status(self.tr("no_default_pdf_app"), error=True)

    def choose_acrobat_folder(self) -> None:
        folder = filedialog.askdirectory(title=self.tr("acrobat_folder_title"))
        if not folder: return
        executable = acrobat_from_folder(folder)
        if executable:
            self.preferences["acrobat_path"] = str(executable)
            save_settings(self.preferences)
            self.show_status(self.tr("acrobat_path_saved"))
            self.choose_acrobat_button.grid_forget()
        else:
            self.show_status(self.tr("acrobat_not_in_folder"), error=True)

    def open_output_folder(self) -> None:
        if not self.result or not self.result.outputs: return
        if not reveal_file(self.result.outputs[0]):
            self.show_status(self.tr("folder_open_failed"), error=True)

    def show_status(self, text: str, error: bool = False) -> None:
        self._set_status_content(text, error=error)
        self.status_frame.grid(row=self.status_row, column=0, pady=5, sticky="w")
        self.after_idle(self._resize_status_to_content)

    def show_ghostscript_status(self, prefix: str) -> None:
        self._set_status_content(prefix, link=True, error=True)
        self.status_frame.grid(row=self.status_row, column=0, pady=5, sticky="w")
        self.after_idle(self._resize_status_to_content)

    def _set_status_content(self, text: str, link: bool = False,
                            error: bool = False) -> None:
        full_text = text + ("Ghostscript." if link else "")
        self._status_min_lines = max(
            1,
            sum(
                max(1, len(textwrap.wrap(line, width=48)))
                for line in full_text.splitlines() or [""]
            ),
        )
        self.status.configure(state="normal")
        self.status.delete("1.0", "end")
        self.status.configure(
            fg=RED if error else self._apply_appearance_mode(TEXT),
            cursor="arrow",
        )
        self.status.insert("end", text)
        if link:
            self.status.insert("end", "Ghostscript", "ghostscript")
            self.status.insert("end", ".")
        self.status.configure(state="disabled", height=self._status_min_lines)

    def _resize_status_to_content(self) -> None:
        if not self.status.winfo_exists() or not self.status.winfo_ismapped():
            return
        self.status.update_idletasks()
        try:
            lines = int(self.status.count("1.0", "end-1c", "displaylines")[0])
        except (tk.TclError, TypeError):
            lines = 1
        self.status.configure(height=max(self._status_min_lines, lines))
        self.after_idle(self.fit_window_to_content)

    def close_app(self) -> None:
        if self.worker and self.worker.is_alive():
            self.show_status(self.tr("wait_for_job"), error=True)
            return
        if "open_result_file" in self.fields:
            self.save_acrobat_preferences()
        self.destroy()


if __name__ == "__main__":
    configure_logging()
    ScanItEasyApp().mainloop()
