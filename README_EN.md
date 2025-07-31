# Offitrans

<div align="center">

A powerful Office file translation tool library that supports batch translation of PDF, Excel, PPT, and Word documents.

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/minglu6/Offitrans.svg)](https://github.com/minglu6/Offitrans/stargazers)

[中文文档](README.md) | **English**

</div>

## ✨ Features

- 🔄 **Multi-format Support**: Supports PDF, Excel, PPT, and Word document translation
- 🌍 **Multi-language Translation**: Supports Chinese, English, Thai, Japanese, Korean, French, German, Spanish, and more
- 🎨 **Format Preservation**: Maintains original formatting, styles, and layout after translation
- 🖼️ **Image Protection**: Complete protection of images in documents without distortion
- ⚡ **Batch Processing**: Efficient batch translation with text deduplication and API call optimization
- 🔧 **Easy Integration**: Clean API design for easy integration into other projects
- 📊 **Rich Text Support**: Supports complex rich text formats and merged cells
- 🛡️ **Error Handling**: Comprehensive error handling and retry mechanisms

## 🚀 Installation

```bash
pip install -r requirements.txt
```

Or install from source:

```bash
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans
pip install -e .
```

## 📦 Dependencies

- `openpyxl` - Excel file processing
- `python-docx` - Word document processing  
- `PyPDF2` - PDF file processing
- `python-pptx` - PowerPoint file processing
- `requests` - HTTP requests
- `Pillow` - Image processing

## 🎯 Quick Start

### Excel File Translation

```python
from excel_translate import ExcelTranslatorV2

# Create translator instance
translator = ExcelTranslatorV2()

# Translate Excel file
success = translator.replace_text_in_excel(
    excel_path="input.xlsx",
    output_path="output_translated.xlsx", 
    target_language="th"  # Translate to Thai
)

if success:
    print("Translation completed!")
```

### PDF File Translation

```python
from pdf_translate import translate_pdf

# Translate PDF file
translate_pdf(
    input_path="document.pdf",
    output_path="document_translated.pdf",
    target_language="en"
)
```

### Word Document Translation

```python
from word_translate import docx_translate

# Translate Word document
docx_translate(
    input_file="document.docx",
    output_file="document_translated.docx",
    target_language="th"
)
```

### PPT File Translation

```python
from ppt_translate import translate_ppt

# Translate PPT file
translate_ppt(
    input_path="presentation.pptx", 
    output_path="presentation_translated.pptx",
    target_language="ja"
)
```

## 🔧 Configuration

### Google Translate API Configuration

The project supports Google Translate API, requiring API key configuration:

```python
from translate_tools import GoogleTranslator

translator = GoogleTranslator(
    source_lang="zh",
    target_lang="en", 
    api_key="your-google-translate-api-key"
)
```

### Supported Language Codes

| Language | Code |
|----------|------|
| Chinese | zh |
| English | en |
| Thai | th |
| Japanese | ja |
| Korean | ko |
| French | fr |
| German | de |
| Spanish | es |

## 📖 Detailed Documentation

### Excel Translator Advanced Features

```python
from excel_translate import ExcelTranslatorV2

# Create translator with custom configuration
translator = ExcelTranslatorV2(
    translate_api_key="your-api-key",
    font_size_adjustment=0.8  # Font size adjustment ratio
)

# Analyze Excel file structure
analysis = translator.analyze_excel_structure("input.xlsx")

# Smart column width adjustment
translator.smart_adjust_column_width("output.xlsx")
```

### Batch Translation Tools

```python
from translate_tools import GoogleTranslator

translator = GoogleTranslator(
    source_lang="zh",
    target_lang="th",
    max_workers=5  # Number of concurrent threads
)

# Batch translate texts
texts = ["Hello", "World", "Translation"]
translated = translator.translate_text_batch(texts)
```

## 🤝 Contributing

We welcome contributions of any kind! Please check [CONTRIBUTING.md](CONTRIBUTING.md) to learn how to participate in project development.

### Development Environment Setup

```bash
# Clone the repository
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run tests
python -m pytest tests/
```

## 📝 License

This project is licensed under the [MIT License](LICENSE).

## 🔗 Related Links

- [GitHub Repository](https://github.com/minglu6/Offitrans)
- [Issue Tracker](https://github.com/minglu6/Offitrans/issues)
- [Releases](https://github.com/minglu6/Offitrans/releases)

## 📊 Project Status

- ✅ Excel Translation - Full support
- ✅ Word Translation - Basic support  
- ✅ PDF Translation - Basic support
- ✅ PPT Translation - Basic support
- 🔄 OCR Support - In development
- 🔄 More Translation Engines - Planned

## 🏗️ Architecture

### Core Components

- **translate_tools/**: Core translation engine and utilities
  - `base.py`: Abstract translator base class
  - `google_translate.py`: Google Translate API implementation
  - `cache.py`: Translation caching mechanism
  - `utils.py`: Common utilities

- **excel_translate/**: Excel-specific translation logic
  - `translate_excel.py`: Excel translator with format preservation

- **word_translate/**: Word document translation
  - `docx_translate.py`: Word document translator

- **pdf_translate/**: PDF translation
  - `translate_pdf.py`: PDF translator

- **ppt_translate/**: PowerPoint translation
  - `translate_ppt.py`: PowerPoint translator

### Key Features

#### Format Preservation
- Maintains font styles, colors, and sizes
- Preserves cell formatting and borders
- Supports merged cells and rich text
- Protects images and charts

#### Performance Optimization
- Text deduplication to reduce API calls
- Concurrent translation processing
- Smart caching mechanisms
- Progress tracking and statistics

#### Error Handling
- Comprehensive retry mechanisms
- Graceful fallback strategies
- Detailed error logging
- Recovery from partial failures

## 🧪 Testing

### Running Tests

```bash
# Run all tests
pytest

# Run specific test file
pytest tests/test_excel_translate.py

# Run tests with coverage
pytest --cov=. --cov-report=html
```

### Test Structure

```
tests/
├── test_excel_translate.py    # Excel translation tests
├── test_google_translate.py   # Google Translate API tests
├── test_base_translator.py    # Base translator tests
└── fixtures/                  # Test data files
    ├── sample.xlsx
    ├── sample.docx
    └── sample.pdf
```

## 🌟 Use Cases

### For Businesses
- **Document Localization**: Translate business documents for international markets
- **Report Translation**: Convert financial reports and presentations
- **Contract Translation**: Translate legal documents while preserving formatting

### For Developers
- **API Integration**: Easy integration into existing applications
- **Batch Processing**: Process large volumes of documents efficiently
- **Custom Workflows**: Build custom translation pipelines

### For Individuals
- **Academic Papers**: Translate research documents and presentations
- **Personal Documents**: Convert personal files between languages
- **Educational Content**: Translate learning materials

## 🙏 Acknowledgments

Thanks to all developers and users who have contributed to this project!

### Special Thanks
- OpenPyXL team for excellent Excel processing capabilities
- Google Translate API for reliable translation services
- Python community for amazing libraries and tools

---

<div align="center">
If this project helps you, please give us a ⭐️
</div>