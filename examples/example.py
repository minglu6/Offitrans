#!/usr/bin/env python3
"""
Offitrans Usage Examples

This file demonstrates how to use Offitrans for translating various Office files.
"""

import os
from offitrans import ExcelProcessor, GoogleTranslator
from offitrans.processors import get_processor_by_extension
from offitrans.core.config import Config


def example_excel_translation():
    """Excel file translation example"""
    print("=" * 50)
    print("Excel File Translation Example")
    print("=" * 50)

    # Create Excel translator
    translator = ExcelProcessor(font_size_adjustment=0.8)  # Font size adjustment ratio

    # Example file paths
    input_file = "example_input.xlsx"
    output_file = "example_output_translated.xlsx"

    # Check if input file exists
    if os.path.exists(input_file):
        print(f"Translating file: {input_file}")

        # Analyze file structure
        print("Analyzing Excel file structure...")
        # Note: structure analysis is now handled internally by the processor

        # Execute translation
        print("Starting translation...")
        success = translator.process_file(
            excel_path=input_file,
            output_path=output_file,
            target_language="en",  # Translate to English
        )

        if success:
            print(f"Translation successful! Output file: {output_file}")

            # Smart column width adjustment
            print("Adjusting column widths...")
            # Note: column width adjustment is now handled automatically by the processor
            print("Column width adjustment completed!")
        else:
            print("Translation failed")
    else:
        print(f"Warning: Input file does not exist: {input_file}")
        print("Please prepare an Excel file for testing")


def example_text_translation():
    """Text translation example"""
    print("\n" + "=" * 50)
    print("Text Translation Example")
    print("=" * 50)

    # Create translator
    translator = GoogleTranslator(source_lang="zh", target_lang="en", max_workers=3)

    # Single text translation
    text = "Hello, world!"
    print(f"Original: {text}")

    translated = translator.translate_text(text)
    print(f"Translated: {translated}")

    # Batch text translation
    texts = [
        "Welcome to Offitrans",
        "This is a powerful translation tool",
        "Supports multiple Office file formats",
        "Maintains original format and style",
    ]

    print(f"\nBatch translating {len(texts)} texts:")
    for i, text in enumerate(texts):
        print(f"{i+1}. {text}")

    print("\nTranslation results:")
    translated_texts = translator.translate_text_batch(texts)
    for i, (original, translated) in enumerate(zip(texts, translated_texts)):
        print(f"{i+1}. {original} -> {translated}")


def example_supported_languages():
    """Supported languages example"""
    print("\n" + "=" * 50)
    print("Supported Languages")
    print("=" * 50)

    from offitrans.translators.google import get_supported_languages

    languages = get_supported_languages()
    print("Currently supported languages:")
    for code, name in languages.items():
        print(f"  {code}: {name}")


def main():
    """Main function"""
    print("Offitrans Usage Examples")
    print("This example demonstrates how to use Offitrans for file translation")

    try:
        # Excel translation example
        example_excel_translation()

        # Text translation example
        example_text_translation()

        # Supported languages
        example_supported_languages()

        print("\n" + "=" * 50)
        print("Examples completed!")
        print("=" * 50)
        print("For more features, please refer to:")
        print("- README.md: Detailed usage documentation")
        print("- CONTRIBUTING.md: Contribution guide")
        print("- GitHub: https://github.com/your-username/Offitrans")

    except ImportError as e:
        print(f"Import error: {e}")
        print("Please ensure all dependencies are properly installed:")
        print("pip install -r requirements.txt")

    except Exception as e:
        print(f"Runtime error: {e}")
        print("Please check configuration and input files")


if __name__ == "__main__":
    main()
