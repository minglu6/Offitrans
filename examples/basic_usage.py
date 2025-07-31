#!/usr/bin/env python3
"""
Basic Usage Examples for Offitrans

This example demonstrates the basic functionality of Offitrans
for translating different types of Office documents.
"""

import os
import logging
from pathlib import Path

# Import the new Offitrans components
from offitrans import ExcelProcessor, GoogleTranslator
from offitrans.processors import get_processor_by_extension
from offitrans.core.config import Config

# Set up logging to see what's happening
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def basic_excel_translation():
    """
    Basic Excel file translation example
    """
    print("=" * 60)
    print("Basic Excel Translation Example")
    print("=" * 60)
    
    # Create a custom configuration
    config = Config()
    config.translator.max_workers = 3  # Use 3 concurrent workers
    config.processor.font_size_adjustment = 0.8  # Adjust font size to 80%
    config.cache.enabled = True  # Enable caching
    
    # Create Excel processor with custom config
    processor = ExcelProcessor(config=config)
    
    # Example file paths (adjust these to your actual files)
    input_file = "examples/sample_files/sample.xlsx"
    output_file = "examples/sample_files/sample_translated.xlsx"
    
    # Check if input file exists
    if not os.path.exists(input_file):
        print(f"Warning: Sample file not found: {input_file}")
        print("Create a sample Excel file with some Chinese text to test")
        return
    
    try:
        # Translate the Excel file
        print(f"Translating Excel file: {input_file}")
        success = processor.process_file(
            input_file, 
            output_file, 
            target_language="en"  # Translate to English
        )
        
        if success:
            print(f"Translation completed successfully!")
            print(f"Output file: {output_file}")
            
            # Show statistics
            stats = processor.get_stats()
            print(f"Statistics:")
            print(f"- Files processed: {stats['total_files_processed']}")
            print(f"- Texts translated: {stats['total_texts_translated']}")
            print(f"- Characters translated: {stats['total_chars_translated']}")
        else:
            print("Translation failed")
            
    except Exception as e:
        print(f"Error during translation: {e}")


def basic_translator_usage():
    """
    Basic translator usage example
    """
    print("\n" + "=" * 60)
    print("Basic Translator Usage Example")
    print("=" * 60)
    
    # Create a Google translator
    translator = GoogleTranslator(
        source_lang="en",  # English
        target_lang="zh",  # Chinese
        use_free_api=False,  # Use free Google Translate API
        api_key="your_api_key_here",  # Replace with your actual API key
        max_workers=2
    )
    
    # Single text translation
    print("Single Text Translation:")
    text = "Hello, world! This is a test."
    print(f"Original: {text}")
    
    try:
        translated = translator.translate_text(text)
        print(f"Translated: {translated}")
    except Exception as e:
        print(f"Translation failed: {e}")
    
    # Batch text translation
    print("\nBatch Text Translation:")
    texts = [
        "Welcome to Offitrans",
        "This is a powerful translation tool",
        "Supports multiple Office file formats",
        "123",  # This should not be translated
        "test@email.com"  # This should not be translated
    ]
    
    print("Original texts:")
    for i, text in enumerate(texts, 1):
        print(f"{i}. {text}")
    
    try:
        translated_texts = translator.translate_text_batch(texts)
        print("\nTranslated texts:")
        for i, (original, translated) in enumerate(zip(texts, translated_texts), 1):
            status = "→" if translated != original else "↷ (skipped)"
            print(f"{i}. {original} {status} {translated}")
    except Exception as e:
        print(f"Batch translation failed: {e}")


def processor_factory_example():
    """
    Example of using processor factory functions
    """
    print("\n" + "=" * 60)
    print("Processor Factory Example")
    print("=" * 60)
    
    # Example file paths
    sample_files = [
        "sample.xlsx",
        "sample.docx", 
        "sample.pdf",
        "sample.pptx"
    ]
    
    for file_path in sample_files:
        try:
            # Get appropriate processor based on file extension
            processor = get_processor_by_extension(file_path)
            print(f"File {file_path} → {processor.__class__.__name__}")
            
            # Show what file types this processor supports
            extensions = []
            for ext in ['.xlsx', '.xls', '.docx', '.doc', '.pdf', '.pptx', '.ppt']:
                if processor.supports_file_type(f"test{ext}"):
                    extensions.append(ext)
            
            print(f"Supports: {', '.join(extensions)}")
            
        except ValueError as e:
            print(f"{file_path} → {e}")
        except Exception as e:
            print(f"{file_path} → Error: {e}")


def translation_with_different_languages():
    """
    Example of translating to different target languages
    """
    print("\n" + "=" * 60)
    print("Multi-Language Translation Example")
    print("=" * 60)
    
    # Sample text in Chinese
    chinese_text = "Hello, world! Welcome to Offitrans translation tool."
    
    # Different target languages
    target_languages = {
        "en": "English",
        "th": "Thai", 
        "ja": "Japanese",
        "ko": "Korean",
        "fr": "French",
        "de": "German",
        "es": "Spanish"
    }
    
    print(f"Original text: {chinese_text}")
    print("\nTranslations to different languages:")
    
    for lang_code, lang_name in target_languages.items():
        try:
            translator = GoogleTranslator(
                source_lang="zh",
                target_lang=lang_code,
                use_free_api=True,
                max_workers=1
            )
            
            translated = translator.translate_text(chinese_text)
            print(f"{lang_name} ({lang_code}): {translated}")
            
        except Exception as e:
            print(f"{lang_name} ({lang_code}): Failed - {e}")


def main():
    """
    Main function to run all examples
    """
    print("Offitrans Basic Usage Examples")
    print("This example demonstrates how to use the Offitrans library")
    
    # Run all examples
    basic_translator_usage()
    # processor_factory_example()
    # translation_with_different_languages()
    # basic_excel_translation()
    
    print("\n" + "=" * 60)
    print("Examples completed!")
    print("=" * 60)
    print("Tips:")
    print("1. Create sample files in examples/sample_files/ to test file translation")
    print("2. Adjust the file paths in the examples to match your files")
    print("3. Check the logs for detailed information about the translation process")
    print("4. Use Config class to customize translator and processor behavior")
    print("\nFor more examples, check:")
    print("- examples/advanced_excel.py")
    print("- examples/batch_processing.py")
    print("- README.md and README_EN.md")


if __name__ == "__main__":
    main()