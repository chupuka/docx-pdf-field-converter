# Docx → PDF Field Converter

Desktop application that converts `.docx` files with underscores into PDF documents with interactive text input fields.

[Русская версия](#rus)

---

## What It Does

1. **Load DOCX files** - add one or more .docx files
2. **Find underscores** - automatically detects `___` patterns (3+ consecutive underscores)
3. **Convert to PDF** - converts document to PDF using Microsoft Word
4. **Add fields** - creates interactive text input fields over each underscore
5. **Save** - saves ready PDFs to selected folder

## Requirements

- Windows / macOS / Linux
- Python 3.8+
- Microsoft Word (required for DOCX → PDF conversion)

## Requirements

### System Requirements

| Platform | Required Software |
|----------|-------------------|
| **Windows** | Microsoft Word (2016 or later recommended) |
| **macOS** | Microsoft Word |
| **Linux** | LibreOffice (latest version) |

### Installation

### Pre-built executable (recommended)
1. Download the latest release from the [Releases](https://github.com/chupuka/docx-pdf-field-converter/releases) page
2. Install the required software (see table above)
3. Run the executable

### From source
```bash
# Clone the repository
git clone https://github.com/chupuka/docx-pdf-field-converter.git
cd docx-pdf-field-converter

# Install dependencies
pip install -r requirements.txt

# Install required software (see table above)

# Run the application
python app.py
```

## Dependencies

- **pymupdf** (PyMuPDF) - PDF processing
- **docx2pdf** - DOCX to PDF conversion via Microsoft Word
- **pyinstaller** - Create .exe (optional)

## Usage

1. Click **"➕ Add Files"** - select one or more .docx files
2. Remove files if needed with **"❌ Remove"**
3. Click **"📂 Select Folder"** - choose where to save PDFs
4. Click **"🚀 Convert to PDF"**
5. Ready PDFs will appear in the selected folder

## Document Requirements

- File must be in `.docx` format
- Underscores must be continuous (minimum 3 characters): `___`, `________`
- Underscores can be in text or tables

## Build .exe

```bash
pyinstaller app.spec --clean
```

The executable will be in the `dist` folder.

## Supported Platforms

| Platform | Works | Notes |
|----------|-------|-------|
| Windows | ✅ | Requires Word |
| macOS | ✅ | Requires Word |
| Linux | ✅ | Requires Word |

## License

MIT

---

<a name="rus"></a>
# Docx → PDF Конвертер полей ввода

Настольное приложение, которое конвертирует `.docx` файлы с подчеркиваниями в PDF с интерактивными текстовыми полями для ввода.

## Что делает приложение

1. **Загрузка DOCX файлов** - добавляете один или несколько файлов .docx
2. **Поиск подчеркиваний** - автоматически находит последовательности `___` (3 и более символа)
3. **Конвертация в PDF** - преобразует документ в PDF через Microsoft Word
4. **Добавление полей** - поверх каждого подчеркивания создается интерактивное текстовое поле
5. **Сохранение** - сохраняет готовые PDF в выбранную папку

## Требования

### Системные требования

| Платформа | Требуемое ПО |
|-----------|--------------|
| **Windows** | Microsoft Word (рекомендуется 2016 и новее) |
| **macOS** | Microsoft Word |
| **Linux** | LibreOffice (последняя версия) |

## Установка

### Готовый исполняемый файл (рекомендуется)
1. Скачайте последний релиз со страницы [Releases](https://github.com/chupuka/docx-pdf-field-converter/releases)
2. Установите требуемое ПО (см. таблицу выше)
3. Запустите исполняемый файл

### Из исходного кода
```bash
git clone https://github.com/chupuka/docx-pdf-field-converter.git
cd docx-pdf-field-converter
pip install -r requirements.txt
# Установите требуемое ПО (см. таблицу выше)
python app.py
```

## Использование

1. Нажмите **"➕ Добавить файлы"** - выберите .docx файлы
2. При необходимости удалите лишние кнопкой **"❌ Удалить"**
3. Нажмите **"📂 Выбрать папку"** - укажите куда сохранить PDF
4. Нажмите **"🚀 Конвертировать в PDF"**
5. Готовые PDF появятся в выбранной папке