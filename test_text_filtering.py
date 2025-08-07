#!/usr/bin/env python3
"""
Test the updated text filtering logic
"""

from offitrans.core.utils import should_translate_text

def test_text_filtering():
    """Test the updated text filtering logic"""
    
    # These are the texts from the Excel file that should be translated
    texts_should_translate = [
        "Name",
        "Country", 
        "Language",
        "Alice",
        "USA",
        "English",
        "Bob", 
        "France",
        "French",
        "Translation Service Test Excel",
        "This Excel file contains English content for translation testing."
    ]
    
    # These are texts that should NOT be translated
    texts_should_not_translate = [
        "123",
        "abc123",
        "ID123",
        "A1B2C3",
        "A",
        "B",
        "=SUM(A1:A2)",
        "www.google.com",
        "test@email.com",
        "12.5%",
        "100px",
        "$50.00",
        "2024-01-15",
        "14:30:00"
    ]
    
    print("=== Testing Updated Text Filtering Logic ===\n")
    
    print("1. Texts that SHOULD be translated:")
    print("-" * 50)
    all_correct = True
    for text in texts_should_translate:
        should_translate = should_translate_text(text)
        status = "✓ PASS" if should_translate else "✗ FAIL"
        print(f"  '{text}' → {should_translate} {status}")
        if not should_translate:
            all_correct = False
    
    print(f"\nResult: {'All tests passed' if all_correct else 'Some tests failed'}")
    
    print("\n2. Texts that should NOT be translated:")
    print("-" * 50)
    all_correct = True
    for text in texts_should_not_translate:
        should_translate = should_translate_text(text)
        status = "✓ PASS" if not should_translate else "✗ FAIL"
        print(f"  '{text}' → {should_translate} {status}")
        if should_translate:
            all_correct = False
    
    print(f"\nResult: {'All tests passed' if all_correct else 'Some tests failed'}")
    
    # Now test the actual Excel file
    print("\n3. Testing actual Excel extraction with new filtering:")
    print("-" * 50)
    
    from offitrans.processors.excel import ExcelProcessor
    from offitrans.core.config import Config
    from offitrans.translators.google import GoogleTranslator
    
    config = Config()
    translator = GoogleTranslator(use_free_api=True)  # Use free API for testing
    processor = ExcelProcessor(translator=translator, config=config)
    
    sample_file = "/root/projects/github/Offitrans/examples/sample_files/sample.xlsx"
    
    # Extract text
    text_data = processor.extract_text(sample_file)
    
    print(f"Extracted {len(text_data)} text entries from Excel:")
    for i, item in enumerate(text_data, 1):
        text = item['text']
        should_translate = should_translate_text(text)
        status = "→ Will translate" if should_translate else "↷ Will skip"
        print(f"  {i:2d}. '{text}' {status}")

if __name__ == "__main__":
    test_text_filtering()