# ScanItEasy

[Русский](#русский) · [English](#english) · [User guide](https://pavelcraft.github.io/ScanItEasy/)

## Русский

ScanItEasy — бесплатное Windows-приложение для конвертации, объединения и обработки DOCX, PDF, PNG и JPG. Документы обрабатываются локально на компьютере.

Текущая версия: **2.0.0**.

Главная особенность приложения — имитация сканирования: электронные страницы документа Word преобразуются так, будто их распечатали на чёрно-белом принтере, сшили лентой и отсканировали. Последнюю, уже подписанную или заверенную часть можно добавить отдельно.

### Возможности

- имитация сканирования прошитого документа;
- DOCX → PDF, PNG или JPG;
- PDF → PNG или JPG;
- объединение любого количества PDF;
- выбор страниц отдельно для каждого PDF;
- PNG/JPG → PDF, включая несколько изображений за один раз;
- перевод PDF в чёрно-белый вид;
- сжатие PDF, PNG и JPG;
- изменение размера изображений с сохранением пропорций.

Поля выбора страниц сохраняют указанный порядок, обратные диапазоны и повторы:

- `1-5` → 1, 2, 3, 4, 5;
- `5-1` → 5, 4, 3, 2, 1;
- `7, 1-3` → 7, 1, 2, 3;
- `3, 1-3` → 3, 1, 2, 3.

Подробное описание: [pavelcraft.github.io/ScanItEasy](https://pavelcraft.github.io/ScanItEasy/).

### Требования

- Windows 10 или 11;
- Microsoft Word — для режимов, работающих с DOCX;
- Python 3.12 — только для запуска из исходного кода;
- Ghostscript — необязателен, но позволяет сильнее сжимать некоторые PDF.

### Запуск из исходного кода

В Git Bash:

```bash
git clone https://github.com/PavelCraft/ScanItEasy.git
cd ScanItEasy
py -3.12 -m venv venv
source venv/Scripts/activate
python -m pip install -r requirements.txt
python ScanItEasy_GUI.py
```

### Компиляция релизного EXE с помощью PyInstaller

Все необходимые изображения, иконки и элементы оформления уже перечислены в
`ScanItEasy_GUI.spec`. Поэтому после активации виртуального окружения запустите
из корня проекта одну команду:

```bash
pyinstaller --clean --noconfirm ScanItEasy_GUI.spec
```

PyInstaller удалит промежуточный кэш предыдущей сборки, включит в приложение
все ресурсы из spec-файла и создаст релизный файл:

```text
dist/ScanItEasy_GUI.exe
```

Папки `build/` и `dist/` генерируются автоматически и не должны добавляться в Git.
Для следующей сборки используется та же самая команда.

<details>
<summary>Полная Bash-команда без использования spec-файла</summary>

```bash
pyinstaller --clean --noconfirm --onefile --noconsole --icon=label.png \
  --add-data "label.png;." \
  --add-data "SIE-no-scan.ico;." \
  --add-data "light.png;." \
  --add-data "dark.png;." \
  --add-data "doc_decorations/Красная лента. Первая страница.png;doc_decorations" \
  --add-data "doc_decorations/Синяя лента. Первая страница.png;doc_decorations" \
  --add-data "doc_decorations/Уголок с красной лентой.png;doc_decorations" \
  --add-data "doc_decorations/Уголок с синей лентой.png;doc_decorations" \
  ScanItEasy_GUI.py
```

</details>

### Ghostscript

ScanItEasy ищет Ghostscript автоматически — вручную указывать путь не требуется.

[Инструкция по установке Ghostscript](https://pavelcraft.github.io/ScanItEasy/ghostscript/).

## English

ScanItEasy is a free Windows application for converting, combining and processing DOCX, PDF, PNG and JPG files. Documents are processed locally on the computer.

Current version: **2.0.0**.

Its main distinguishing feature is scan simulation: electronic Word pages are transformed to look as though they were printed in black and white, bound with ribbon and scanned. A final signed or certified part can be added separately.

### Features

- simulation of a scanned, bound document;
- DOCX → PDF, PNG or JPG;
- PDF → PNG or JPG;
- merging any number of PDFs;
- independent page selection for every PDF;
- PNG/JPG → PDF, including multiple images at once;
- grayscale PDF conversion;
- PDF, PNG and JPG compression;
- proportional image resizing.

Page-selection fields preserve order, reverse ranges and repeated pages:

- `1-5` → 1, 2, 3, 4, 5;
- `5-1` → 5, 4, 3, 2, 1;
- `7, 1-3` → 7, 1, 2, 3;
- `3, 1-3` → 3, 1, 2, 3.

Read the full guide at [pavelcraft.github.io/ScanItEasy/en/](https://pavelcraft.github.io/ScanItEasy/en/).

### Requirements

- Windows 10 or 11;
- Microsoft Word for DOCX operations;
- Python 3.12 only when running from source;
- optional Ghostscript for stronger compression of some PDFs.

### Run from source

In Git Bash:

```bash
git clone https://github.com/PavelCraft/ScanItEasy.git
cd ScanItEasy
py -3.12 -m venv venv
source venv/Scripts/activate
python -m pip install -r requirements.txt
python ScanItEasy_GUI.py
```

### Compile the release EXE with PyInstaller

All required images, icons and decoration assets are already listed in
`ScanItEasy_GUI.spec`. Activate the virtual environment and run this command
from the repository root:

```bash
pyinstaller --clean --noconfirm ScanItEasy_GUI.spec
```

The release executable will be created at:

```text
dist/ScanItEasy_GUI.exe
```

The generated `build/` and `dist/` directories are ignored by Git. Use the same
command for subsequent builds.

<details>
<summary>Full Bash command without the spec file</summary>

```bash
pyinstaller --clean --noconfirm --onefile --noconsole --icon=label.png \
  --add-data "label.png;." \
  --add-data "SIE-no-scan.ico;." \
  --add-data "light.png;." \
  --add-data "dark.png;." \
  --add-data "doc_decorations/Красная лента. Первая страница.png;doc_decorations" \
  --add-data "doc_decorations/Синяя лента. Первая страница.png;doc_decorations" \
  --add-data "doc_decorations/Уголок с красной лентой.png;doc_decorations" \
  --add-data "doc_decorations/Уголок с синей лентой.png;doc_decorations" \
  ScanItEasy_GUI.py
```

</details>

ScanItEasy detects Ghostscript automatically.

[Ghostscript installation guide](https://pavelcraft.github.io/ScanItEasy/en/ghostscript/).

## License

BSD 3-Clause. See [LICENSE](LICENSE).

## Contact

- Email: contact@formacode.dev
- Telegram: [@PavelCraftRobot](https://t.me/PavelCraftRobot)
