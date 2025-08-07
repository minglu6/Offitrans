#!/usr/bin/env python3
"""
Test full translation with the updated filtering logic
"""

from offitrans.processors.excel import ExcelProcessor
from offitrans.core.config import Config
from offitrans.translators.google import GoogleTranslator
import logging

logging.basicConfig(level=logging.INFO)

def test_full_translation():
    """Test full Excel translation"""
    print("=== Testing Full Excel Translation ===\n")
    
    # Setup
    config = Config()
    config.cache.enabled = False  # Disable cache for testing
    translator = GoogleTranslator(use_free_api=True, max_workers=1)
    processor = ExcelProcessor(translator=translator, config=config)
    
    # Files
    input_file = "/root/projects/github/Offitrans/examples/sample_files/sample.xlsx"
    output_file = "/root/projects/github/Offitrans/test_output.xlsx"
    
    try:
        # Test translation to Chinese
        print("1. Translating Excel file to Chinese...")
        print(f"Input: {input_file}")
        print(f"Output: {output_file}")
        
        success = processor.translate_and_save(input_file, output_file, target_language="zh")
        
        if success:
            print("✅ Translation completed successfully!")
            
            # Show statistics
            stats = processor.get_stats()
            print(f"\nStatistics:")
            print(f"- Files processed: {stats.get('total_files_processed', 0)}")
            print(f"- Texts translated: {stats.get('total_texts_translated', 0)}")
            print(f"- Characters translated: {stats.get('total_chars_translated', 0)}")
            print(f"- Translation time: {stats.get('total_translation_time', 0):.2f}s")
        else:
            print("❌ Translation failed")
            return False
        
        # Test with different target language (Thai)
        print("\n2. Testing with Thai translation...")
        output_file_th = "/root/projects/github/Offitrans/test_output_th.xlsx"
        
        success_th = processor.translate_and_save(input_file, output_file_th, target_language="th")
        
        if success_th:
            print("✅ Thai translation completed successfully!")
        else:
            print("❌ Thai translation failed")
        
        return success and success_th
        
    except Exception as e:
        print(f"❌ Error during translation: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    result = test_full_translation()
    print(f"\n{'='*50}")
    print(f"Final result: {'SUCCESS' if result else 'FAILED'}")
    print(f"{'='*50}")