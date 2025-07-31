# API Documentation

This document provides detailed API documentation for Offitrans.

## Core Components

### BaseTranslator

The abstract base class for all translators.

```python
from offitrans.core.base import BaseTranslator

class BaseTranslator(ABC):
    def __init__(self, source_lang="auto", target_lang="en", max_workers=5, **kwargs):
        """
        Initialize the base translator.
        
        Args:
            source_lang: Source language code (default: "auto")
            target_lang: Target language code (default: "en")
            max_workers: Maximum number of concurrent workers (default: 5)
            **kwargs: Additional keyword arguments
        """
    
    @abstractmethod
    def translate_text(self, text: str) -> str:
        """Translate a single text string."""
        
    def translate_text_batch(self, texts: List[str]) -> List[str]:
        """Translate a batch of text strings."""
        
    def get_stats(self) -> Dict[str, int]:
        """Get translation statistics."""
```

### GoogleTranslator

Google Translate API implementation.

```python
from offitrans.translators import GoogleTranslator

translator = GoogleTranslator(
    source_lang="zh",
    target_lang="en",
    api_key=None,  # Optional for free API
    use_free_api=True,  # Use free vs paid API
    max_workers=5
)

# Single translation
result = translator.translate_text("你好世界")

# Batch translation
results = translator.translate_text_batch(["你好", "世界"])

# Language detection
language = translator.detect_language("Hello world")
```

### Configuration

The Config class manages all settings.

```python
from offitrans.core.config import Config

config = Config()

# Translator settings
config.translator.max_workers = 10
config.translator.timeout = 120
config.translator.retry_count = 3

# Cache settings
config.cache.enabled = True
config.cache.cache_file = "custom_cache.json"

# Processor settings
config.processor.preserve_formatting = True
config.processor.image_protection = True
config.processor.font_size_adjustment = 0.8

# Save configuration
config.save_to_file("config.json")

# Load configuration
config = Config("config.json")
```

## File Processors

### BaseProcessor

The abstract base class for all file processors.

```python
from offitrans.processors.base import BaseProcessor

class BaseProcessor(ABC):
    def __init__(self, translator=None, config=None, **kwargs):
        """
        Initialize the base processor.
        
        Args:
            translator: Translator instance to use
            config: Configuration instance
            **kwargs: Additional keyword arguments
        """
    
    @abstractmethod
    def extract_text(self, file_path: str) -> List[Dict[str, Any]]:
        """Extract text content from the file."""
        
    @abstractmethod
    def translate_and_save(self, file_path: str, output_path: str, target_language: str = "en") -> bool:
        """Translate the file and save to output path."""
        
    def process_file(self, input_path: str, output_path: str, target_language: str = "en") -> bool:
        """High-level method to process a file."""
```

### ExcelProcessor

Excel file processor with advanced features.

```python
from offitrans import ExcelProcessor

processor = ExcelProcessor(
    translator=translator,  # Optional custom translator
    config=config,  # Optional custom config
    preserve_formatting=True,
    image_protection=True,
    font_size_adjustment=0.8
)

# Process a single Excel file
success = processor.process_file(
    "input.xlsx",
    "output.xlsx", 
    target_language="en"
)

# Extract text for inspection
text_data = processor.extract_text("input.xlsx")

# Get processing statistics
stats = processor.get_stats()
```

### WordProcessor

Word document processor.

```python
from offitrans.processors import WordProcessor

processor = WordProcessor()

# Process Word document
success = processor.process_file("document.docx", "translated.docx", "en")

# Backward compatibility function
from offitrans.processors.word import docx_translate
success = docx_translate("input.docx", "output.docx", "en")
```

### PDFProcessor

PDF file processor (basic implementation).

```python
from offitrans.processors import PDFProcessor

processor = PDFProcessor()

# Process PDF (outputs as text file currently)
success = processor.process_file("document.pdf", "translated.txt", "en")

# Backward compatibility function
from offitrans.processors.pdf import translate_pdf
success = translate_pdf("input.pdf", "output.txt", "en")
```

### PowerPointProcessor

PowerPoint presentation processor.

```python
from offitrans.processors import PowerPointProcessor

processor = PowerPointProcessor()

# Process PowerPoint
success = processor.process_file("presentation.pptx", "translated.pptx", "en")

# Backward compatibility function
from offitrans.processors.powerpoint import translate_ppt
success = translate_ppt("input.pptx", "output.pptx", "en")
```

## Utility Functions

### Processor Factory

Get processors dynamically based on file type.

```python
from offitrans.processors import get_processor, get_processor_by_extension

# Get by type
excel_processor = get_processor("excel")
pdf_processor = get_processor("pdf")

# Get by file extension
processor = get_processor_by_extension("document.xlsx")
success = processor.process_file("document.xlsx", "translated.xlsx")
```

### Text Utilities

Various text processing utilities.

```python
from offitrans.core.utils import (
    detect_language,
    validate_language_code,
    clean_text,
    should_translate_text,
    filter_translatable_texts,
    deduplicate_texts
)

# Language detection
lang = detect_language("Hello world")  # Returns "en"

# Text filtering
translatable, non_translatable = filter_translatable_texts([
    "Hello world",  # Translatable
    "123",          # Not translatable
    "你好"          # Translatable
])

# Text deduplication
unique_texts, mapping = deduplicate_texts(["hello", "world", "hello"])
```

### Cache Management

Translation cache for improved performance.

```python
from offitrans.core.cache import TranslationCache, get_global_cache

# Create custom cache
cache = TranslationCache("my_cache.json")

# Set and get translations
cache.set("hello", "你好", "en", "zh")
translation = cache.get("hello", "en", "zh")

# Batch operations
translations = {"hello": "你好", "world": "世界"}
cache.set_batch(translations, "en", "zh")

results = cache.get_batch(["hello", "world"], "en", "zh")

# Use global cache
global_cache = get_global_cache()
```

## Exception Handling

Offitrans provides specific exceptions for different error types.

```python
from offitrans.exceptions import (
    OffitransError,
    TranslationError,
    ProcessorError,
    ConfigError,
    ExcelProcessorError,
    WordProcessorError,
    PDFProcessorError,
    PowerPointProcessorError
)

try:
    processor = ExcelProcessor()
    success = processor.process_file("input.xlsx", "output.xlsx")
except ExcelProcessorError as e:
    print(f"Excel processing error: {e}")
    print(f"File: {e.file_path}")
    print(f"Details: {e.details}")
except ProcessorError as e:
    print(f"General processor error: {e}")
except OffitransError as e:
    print(f"Offitrans error: {e}")
```

## Language Codes

Supported language codes:

| Code | Language |
|------|----------|
| `zh` | Chinese |
| `en` | English |
| `th` | Thai |
| `ja` | Japanese |
| `ko` | Korean |
| `fr` | French |
| `de` | German |
| `es` | Spanish |
| `ar` | Arabic |
| `ru` | Russian |
| `auto` | Auto-detect |

## Performance Tips

1. **Use batch translation** for better performance:
   ```python
   # Good
   results = translator.translate_text_batch(texts)
   
   # Avoid
   results = [translator.translate_text(text) for text in texts]
   ```

2. **Enable caching** for repeated translations:
   ```python
   config = Config()
   config.cache.enabled = True
   ```

3. **Adjust max_workers** based on your use case:
   ```python
   # For I/O intensive tasks (API calls)
   translator = GoogleTranslator(max_workers=10)
   
   # For CPU intensive tasks
   translator = GoogleTranslator(max_workers=4)
   ```

4. **Use appropriate font size adjustment** for target languages:
   ```python
   # For languages with larger character sets (Thai, Chinese)
   config.processor.font_size_adjustment = 0.8
   
   # For languages with similar character sizes
   config.processor.font_size_adjustment = 1.0
   ```

## Error Handling Best Practices

1. **Always handle specific exceptions**:
   ```python
   try:
       result = processor.process_file(input_file, output_file)
   except ProcessorError as e:
       # Handle processing-specific errors
       logger.error(f"Processing failed: {e}")
   except TranslationError as e:
       # Handle translation-specific errors
       logger.error(f"Translation failed: {e}")
   ```

2. **Check return values**:
   ```python
   success = processor.process_file(input_file, output_file)
   if not success:
       # Handle failure case
       stats = processor.get_stats()
       logger.error(f"Failed files: {stats['failed_files']}")
   ```

3. **Monitor statistics**:
   ```python
   stats = processor.get_stats()
   logger.info(f"Processed {stats['successful_files']} files successfully")
   logger.info(f"Failed {stats['failed_files']} files")
   ```