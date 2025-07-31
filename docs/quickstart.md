# Quick Start Guide

Get up and running with Offitrans in minutes! This guide will walk you through the basic usage of the library.

## Installation

First, install Offitrans and its dependencies:

```bash
pip install offitrans

# For Excel support
pip install openpyxl

# For Word support  
pip install python-docx

# For PDF support
pip install PyPDF2

# For PowerPoint support
pip install python-pptx
```

## Basic Usage

### 1. Simple Text Translation

```python
from offitrans import GoogleTranslator

# Create a translator
translator = GoogleTranslator(
    source_lang="zh",  # Chinese
    target_lang="en"   # English
)

# Translate single text
result = translator.translate_text("你好世界")
print(result)  # Output: Hello world

# Translate multiple texts
texts = ["你好", "世界", "欢迎"]
results = translator.translate_text_batch(texts)
print(results)  # Output: ['Hello', 'World', 'Welcome']
```

### 2. Excel File Translation

```python
from offitrans import ExcelProcessor

# Create processor
processor = ExcelProcessor()

# Translate Excel file
success = processor.process_file(
    input_path="chinese_file.xlsx",
    output_path="english_file.xlsx",
    target_language="en"
)

if success:
    print("Excel file translated successfully!")
```

### 3. Word Document Translation

```python
from offitrans.processors import WordProcessor

# Create processor
processor = WordProcessor()

# Translate Word document
success = processor.process_file(
    input_path="document.docx",
    output_path="translated_document.docx",
    target_language="en"
)
```

### 4. Auto-detect File Type

```python
from offitrans.processors import get_processor_by_extension

# Automatically get the right processor for your file
processor = get_processor_by_extension("my_file.xlsx")

# Translate the file
success = processor.process_file(
    "my_file.xlsx", 
    "translated_file.xlsx", 
    "en"
)
```

## Advanced Configuration

### Custom Configuration

```python
from offitrans.core.config import Config
from offitrans import ExcelProcessor

# Create custom configuration
config = Config()
config.translator.max_workers = 3
config.translator.timeout = 30
config.cache.enabled = True
config.processor.preserve_formatting = True
config.processor.image_protection = True

# Use with processor
processor = ExcelProcessor(config=config)
```

### Different Target Languages

```python
from offitrans import GoogleTranslator

# Translate to different languages
languages = {
    "en": "English",
    "th": "Thai", 
    "ja": "Japanese",
    "fr": "French"
}

text = "你好世界"

for lang_code, lang_name in languages.items():
    translator = GoogleTranslator(
        source_lang="zh",
        target_lang=lang_code
    )
    result = translator.translate_text(text)
    print(f"{lang_name}: {result}")
```

## Working with Different File Types

### Excel Files with Rich Formatting

```python
from offitrans import ExcelProcessor
from offitrans.core.config import Config

# Configure for rich formatting preservation
config = Config()
config.processor.preserve_formatting = True
config.processor.image_protection = True
config.processor.font_size_adjustment = 0.8  # Adjust for target language

processor = ExcelProcessor(config=config)

success = processor.process_file(
    "formatted_spreadsheet.xlsx",
    "translated_spreadsheet.xlsx", 
    "th"  # Thai
)
```

### Batch Processing Multiple Files

```python
from offitrans.processors import get_processor_by_extension
import os

# List of files to translate
files_to_translate = [
    ("report1.xlsx", "report1_en.xlsx"),
    ("presentation.pptx", "presentation_en.pptx"),
    ("document.docx", "document_en.docx")
]

# Process each file
for input_file, output_file in files_to_translate:
    if os.path.exists(input_file):
        processor = get_processor_by_extension(input_file)
        success = processor.process_file(input_file, output_file, "en")
        
        if success:
            print(f"✅ {input_file} -> {output_file}")
        else:
            print(f"❌ Failed: {input_file}")
    else:
        print(f"⚠️  File not found: {input_file}")
```

## Error Handling

```python
from offitrans import ExcelProcessor
from offitrans.exceptions import ExcelProcessorError, TranslationError

processor = ExcelProcessor()

try:
    success = processor.process_file("input.xlsx", "output.xlsx", "en")
    
    if success:
        # Get statistics
        stats = processor.get_stats()
        print(f"Translated {stats['total_texts_translated']} texts")
    else:
        print("Translation failed")
        
except ExcelProcessorError as e:
    print(f"Excel processing error: {e}")
    if e.file_path:
        print(f"Problem file: {e.file_path}")
        
except TranslationError as e:
    print(f"Translation service error: {e}")
    
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Performance Tips

### 1. Enable Caching

```python
from offitrans.core.config import Config

config = Config()
config.cache.enabled = True
config.cache.cache_file = "my_translations.json"

# Translations will be cached and reused
processor = ExcelProcessor(config=config)
```

### 2. Adjust Concurrency

```python
config = Config()

# For many small texts
config.translator.max_workers = 10

# For large files or limited resources
config.translator.max_workers = 2

processor = ExcelProcessor(config=config)
```

### 3. Filter Unnecessary Content

The library automatically filters out content that doesn't need translation:
- Numbers (123, 45.67)
- Email addresses (user@example.com)
- URLs (https://example.com)
- Formulas (=SUM(A1:B2))
- Short codes (ID123, ABC-DEF)

## Real-World Examples

### Corporate Report Translation

```python
from offitrans import ExcelProcessor
from offitrans.core.config import Config

# Configure for business documents
config = Config()
config.processor.preserve_formatting = True
config.processor.image_protection = True
config.cache.enabled = True

# Create processor with custom translator settings
processor = ExcelProcessor(config=config)

# Translate quarterly report
success = processor.process_file(
    "Q4_Report_CN.xlsx",
    "Q4_Report_EN.xlsx",
    "en"
)

if success:
    stats = processor.get_stats()
    print(f"Translated corporate report successfully!")
    print(f"Processed {stats['total_texts_translated']} text elements")
```

### Educational Content Translation

```python
from offitrans.processors import get_processor_by_extension

# Educational materials in different formats
materials = [
    ("lesson_plan.docx", "lesson_plan_thai.docx", "th"),
    ("worksheet.xlsx", "worksheet_english.xlsx", "en"),
    ("presentation.pptx", "presentation_japanese.pptx", "ja")
]

for input_file, output_file, lang in materials:
    processor = get_processor_by_extension(input_file)
    
    success = processor.process_file(input_file, output_file, lang)
    
    if success:
        print(f"✅ Translated {input_file} to {lang}")
    else:
        print(f"❌ Failed to translate {input_file}")
```

## What's Next?

Now that you've got the basics down, explore:

1. **[Advanced Examples](../examples/)** - More complex use cases
2. **[API Documentation](api.md)** - Detailed API reference
3. **[Configuration Guide](configuration.md)** - Advanced configuration options
4. **[Contributing](../CONTRIBUTING.md)** - How to contribute to the project

## Getting Help

If you run into issues:

1. Check the [Installation Guide](installation.md) for common problems
2. Look at the [examples](../examples/) for similar use cases
3. Search [GitHub Issues](https://github.com/minglu6/Offitrans/issues)
4. Create a new issue with details about your problem

Happy translating! 🚀