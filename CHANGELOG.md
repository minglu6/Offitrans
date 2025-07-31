# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Initial release preparation
- Comprehensive documentation and examples

## [0.2.0] - 2024-01-XX

### Added
- Complete project restructuring with modular architecture
- New `offitrans` package with organized modules:
  - `core/` - Base classes, configuration, caching, and utilities
  - `translators/` - Translation engine implementations
  - `processors/` - File format processors
  - `exceptions/` - Custom exception hierarchy
- Enhanced configuration system with file-based and environment variable support
- Comprehensive caching system for improved performance
- Advanced Excel processor with:
  - Rich text formatting preservation
  - Image protection capabilities
  - Smart column width adjustment
  - Font size adjustment for target languages
- Improved error handling and exception hierarchy
- Extensive test suite with unit and integration tests
- Complete documentation including:
  - API documentation
  - Installation guide
  - Quick start guide
  - Advanced usage examples
- CLI interface for command-line usage
- Support for multiple target languages (English, Thai, Japanese, Korean, French, German, Spanish)
- Batch processing capabilities with parallel execution
- Statistics and monitoring features

### Changed
- **BREAKING**: Reorganized import paths (backward compatibility maintained through aliases)
- Enhanced translator interface with better error handling
- Improved Google Translate integration with both free and paid API support
- Better text filtering to avoid translating non-translatable content
- More robust file processing with better error recovery

### Fixed
- Various stability improvements
- Better handling of special characters and formatting
- Improved memory usage for large files

### Dependencies
- Added support for Python 3.7+
- Updated minimum dependency versions
- Added optional dependencies for different file formats

## [0.1.0] - 2023-XX-XX

### Added
- Initial implementation of office file translation
- Basic Excel file translation support
- Word document translation support
- PDF text extraction and translation
- PowerPoint presentation translation
- Google Translate API integration
- Basic error handling
- Simple configuration system

### Features
- Translation of Excel (.xlsx) files with formatting preservation
- Translation of Word (.docx) documents
- Translation of PowerPoint (.pptx) presentations
- Basic PDF text extraction and translation
- Support for Chinese to English translation
- Concurrent translation processing
- Basic caching mechanism

### Dependencies
- Python 3.6+ support
- Core dependencies: requests
- Optional dependencies: openpyxl, python-docx, PyPDF2, python-pptx

---

## Release Notes

### Version 0.2.0 Highlights

This is a major release that introduces a complete architectural overhaul of the Offitrans library. The new version provides:

#### 🏗️ **Modular Architecture**
- Clean separation of concerns with dedicated modules
- Abstract base classes for extensibility
- Plugin-style processor system

#### ⚡ **Enhanced Performance**
- Intelligent caching system
- Parallel processing capabilities
- Memory-efficient file handling

#### 🎨 **Advanced Excel Features**
- Rich text formatting preservation
- Image protection and handling
- Smart column width adjustment
- Font size optimization for different languages

#### 🔧 **Developer Experience**
- Comprehensive API documentation
- Extensive example collection
- Type hints throughout the codebase
- Pre-commit hooks and development tools

#### 🌍 **Multi-language Support**
- Support for 8+ target languages
- Language detection capabilities
- Cultural-aware text processing

#### 🛡️ **Reliability**
- Robust error handling
- Custom exception hierarchy
- Comprehensive test coverage
- Better logging and monitoring

### Migration Guide

For users upgrading from v0.1.x to v0.2.0:

#### Import Changes
```python
# Old (still works via backward compatibility)
from excel_translate import ExcelTranslatorV2

# New (recommended)
from offitrans import ExcelProcessor
```

#### Configuration Changes
```python
# Old
translator = ExcelTranslatorV2(max_workers=5)

# New
from offitrans.core.config import Config
config = Config()
config.translator.max_workers = 5
processor = ExcelProcessor(config=config)
```

#### API Changes
Most existing code will continue to work due to backward compatibility aliases, but we recommend migrating to the new API for better features and performance.

### Breaking Changes

- **Import paths**: While backward compatibility is maintained, the recommended import paths have changed
- **Configuration system**: New centralized configuration system (old parameters still work)
- **Exception types**: New exception hierarchy (old exceptions are still raised for compatibility)

### Deprecations

- Direct parameter passing to processors (use Config instead)
- Old import paths (will be removed in v1.0.0)
- Legacy function names (use new processor classes)

### Support

- **Python versions**: 3.7+ (dropped support for 3.6)
- **Dependencies**: Updated minimum versions for security and compatibility
- **Platforms**: Windows, macOS, Linux

For detailed migration instructions and new features, see the [documentation](docs/)