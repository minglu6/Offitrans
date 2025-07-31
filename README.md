# Offitrans

<div align="center">

A powerful Office file translation tool library that supports batch translation of PDF, Excel, PPT, and Word documents.

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PyPI](https://img.shields.io/pypi/v/offitrans.svg)](https://pypi.org/project/offitrans/)
[![GitHub Stars](https://img.shields.io/github/stars/minglu6/Offitrans.svg)](https://github.com/minglu6/Offitrans/stargazers)

**English** | [中文](README_ZH.md)

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
- 💻 **Command Line Tool**: Convenient CLI interface with batch operation support

## 🚀 Installation

### Install from PyPI (Recommended)

```bash
# Basic version
pip install offitrans

# Full version with all dependencies
pip install offitrans[full]

# Install specific format support as needed
pip install offitrans[excel]      # Excel support
pip install offitrans[word]       # Word support  
pip install offitrans[pdf]        # PDF support
pip install offitrans[powerpoint] # PowerPoint support
```

### Install from Source

```bash
git clone https://github.com/minglu6/Offitrans.git
cd Offitrans
pip install -e .
```

## 🎯 Quick Start

### Command Line Usage

```bash
# Translate Excel file
offitrans input.xlsx -t zh -o output_zh.xlsx

# Translate PDF file
offitrans document.pdf -t en -o document_en.pdf

# Translate Word document
offitrans document.docx -t th -o document_th.docx

# Translate PowerPoint presentation
offitrans presentation.pptx -t ja -o presentation_ja.pptx

# View all options
offitrans --help
```

### Python API Usage

#### Excel File Translation

```python
from offitrans.processors.excel import ExcelProcessor
from offitrans.translators.google import GoogleTranslator

# Create translator and processor
translator = GoogleTranslator()
processor = ExcelProcessor()

# Translate Excel file
processor.translate_file(
    input_path="input.xlsx",
    output_path="output_translated.xlsx",
    translator=translator,
    target_lang="th"  # Translate to Thai
)
```

#### PDF File Translation

```python
from offitrans.processors.pdf import PDFProcessor
from offitrans.translators.google import GoogleTranslator

translator = GoogleTranslator()
processor = PDFProcessor()

processor.translate_file(
    input_path="document.pdf",
    output_path="document_translated.pdf",
    translator=translator,
    target_lang="en"
)
```

#### Word Document Translation

```python
from offitrans.processors.word import WordProcessor
from offitrans.translators.google import GoogleTranslator

translator = GoogleTranslator()
processor = WordProcessor()

processor.translate_file(
    input_path="document.docx",
    output_path="document_translated.docx",
    translator=translator,
    target_lang="th"
)
```

#### PowerPoint Translation

```python
from offitrans.processors.powerpoint import PowerPointProcessor
from offitrans.translators.google import GoogleTranslator

translator = GoogleTranslator()
processor = PowerPointProcessor()

processor.translate_file(
    input_path="presentation.pptx",
    output_path="presentation_translated.pptx",
    translator=translator,
    target_lang="ja"
)
```

## 🔧 Configuration

### Google Translate API Configuration

```python
from offitrans.translators.google import GoogleTranslator

# Using API key
translator = GoogleTranslator(api_key="your-google-translate-api-key")

# Or set environment variable
import os
os.environ['GOOGLE_TRANSLATE_API_KEY'] = 'your-api-key'
translator = GoogleTranslator()
```

### Supported Language Codes

| Language | Code | Language | Code |
|----------|------|----------|------|
| Chinese | zh | English | en |
| Thai | th | Japanese | ja |
| Korean | ko | French | fr |
| German | de | Spanish | es |

## 📖 Advanced Usage

### Batch Translation

```python
from offitrans.core.utils import BatchProcessor
from offitrans.translators.google import GoogleTranslator

# Batch process multiple files
processor = BatchProcessor()
translator = GoogleTranslator()

files = ["doc1.xlsx", "doc2.docx", "doc3.pdf"]
processor.process_files(files, translator, target_lang="en")
```

### Custom Translator

```python
from offitrans.translators.base_api import BaseTranslator

class CustomTranslator(BaseTranslator):
    def translate(self, text, source_lang="auto", target_lang="en"):
        # Implement custom translation logic
        return translated_text

# Use custom translator
translator = CustomTranslator()
```

### Cache Configuration

```python
from offitrans.core.cache import TranslationCache

# Enable translation cache
cache = TranslationCache(cache_dir="./translation_cache")
translator = GoogleTranslator(cache=cache)
```

## 🏗️ Project Architecture

```
offitrans/
├── cli/                    # Command line interface
│   ├── __init__.py
│   └── main.py            # CLI entry point
├── core/                  # Core functionality
│   ├── base.py           # Base class definitions
│   ├── cache.py          # Caching mechanism
│   ├── config.py         # Configuration management
│   └── utils.py          # Utility functions
├── processors/           # Document processors
│   ├── base.py          # Processor base class
│   ├── excel.py         # Excel processor
│   ├── pdf.py           # PDF processor
│   ├── powerpoint.py    # PowerPoint processor
│   └── word.py          # Word processor
├── translators/         # Translation engines
│   ├── base_api.py      # Translator base class
│   └── google.py        # Google Translate implementation
└── exceptions/          # Exception definitions
    └── errors.py        # Custom exceptions
```

## 🧪 Testing

### Running Tests

```bash
# Install development dependencies
pip install -e .[dev]

# Run all tests
pytest

# Run specific tests
pytest tests/unit/test_processors.py

# Run tests with coverage report
pytest --cov=offitrans --cov-report=html
```

### Test Structure

```
tests/
├── unit/                    # Unit tests
│   ├── test_core.py
│   ├── test_processors.py
│   └── test_translators.py
├── integration/             # Integration tests
│   └── test_end_to_end.py
└── fixtures/               # Test data
    ├── sample.xlsx
    ├── sample.docx
    └── sample.pdf
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

# Install development dependencies
pip install -e .[dev]

# Install pre-commit hooks
pre-commit install

# Run tests
pytest
```

## 📝 License

This project is licensed under the [MIT License](LICENSE).

## 🔗 Related Links

- [GitHub Repository](https://github.com/minglu6/Offitrans)
- [PyPI Package](https://pypi.org/project/offitrans/)
- [Issue Tracker](https://github.com/minglu6/Offitrans/issues)
- [Releases](https://github.com/minglu6/Offitrans/releases)
- [Changelog](CHANGELOG.md)

## 📊 Project Status

- ✅ Excel Translation - Full support
- ✅ Word Translation - Basic support  
- ✅ PDF Translation - Basic support
- ✅ PPT Translation - Basic support
- ✅ CLI Tool - Full support
- 🔄 OCR Support - In development
- 🔄 More Translation Engines - Planned

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

- Google Translate API for reliable translation services
- OpenPyXL team for excellent Excel processing capabilities
- Python community for amazing libraries and tools

---

<div align="center">
If this project helps you, please give us a ⭐️
</div>