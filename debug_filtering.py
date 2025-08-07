#!/usr/bin/env python3
"""
Debug text filtering step by step
"""

import re

def debug_should_translate_text(text: str) -> bool:
    """
    Debug version of should_translate_text with detailed output
    """
    print(f"\nDebugging text: '{text}'")
    
    if not text or not text.strip():
        print("  ❌ Empty or whitespace text")
        return False
    
    text = text.strip()
    print(f"  ✓ After strip: '{text}'")
    
    # Skip pure numbers
    if text.isdigit():
        print("  ❌ Pure digits")
        return False
    
    # Skip pure symbols
    if re.fullmatch(r"[\W_]+", text):
        print("  ❌ Pure symbols")
        return False
    
    # Skip very short pure English letters (like single letters or obvious codes)
    if re.fullmatch(r"[a-zA-Z]+", text) and len(text) <= 2:
        print(f"  ❌ Very short pure letters (len={len(text)})")
        return False
    
    # Skip obvious alphanumeric codes (mixed letters and numbers)
    if re.fullmatch(r"[a-zA-Z0-9]+", text) and re.search(r"\d", text) and re.search(r"[a-zA-Z]", text):
        print("  ❌ Alphanumeric codes")
        return False
    
    # Skip numbers with symbols (prices, percentages, measurements)
    if re.fullmatch(r"[\d\W_]+", text):
        print("  ❌ Numbers with symbols")
        return False
    
    # Skip URLs and emails
    if re.search(r"https?://|www\.|@.*\.|\.com|\.org|\.net|\.edu", text.lower()):
        print("  ❌ URLs and emails")
        return False
    
    # Skip file paths
    if re.search(r"[A-Za-z]:\\|/[a-zA-Z]|\.exe|\.dll|\.pdf|\.docx?|\.xlsx?|\.pptx?", text):
        print("  ❌ File paths")
        return False
    
    # Skip programming identifiers (underscore or camelCase)
    if re.search(r"[a-zA-Z]+_[a-zA-Z]+|[a-z]+[A-Z][a-z]*", text):
        print("  ❌ Programming identifiers")
        return False
    
    # Skip measurements and units
    if re.fullmatch(r"\d+\s*(mm|cm|m|km|kg|g|ml|l|°C|°F|%|px|pt|em|rem|in|ft)", text, re.IGNORECASE):
        print("  ❌ Measurements and units")
        return False
    
    # Skip version numbers
    if re.search(r"v\d+\.\d+|ver\.\d+|version\s*\d+", text.lower()):
        print("  ❌ Version numbers")
        return False
    
    # Skip date formats
    if re.search(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4}", text):
        print("  ❌ Date formats")
        return False
    
    # Skip time formats
    if re.search(r"\d{1,2}:\d{2}(\s*(AM|PM))?", text.upper()):
        print("  ❌ Time formats")
        return False
    
    # Skip formulas (starting with =)
    if text.startswith("="):
        print("  ❌ Formulas")
        return False
    
    # Translate if contains Chinese characters
    if re.search(r"[\u4e00-\u9fff]", text):
        print("  ✅ Contains Chinese characters")
        return True
    
    # Translate if contains other non-ASCII characters (except symbols)
    if re.search(r"[^\x00-\x7f]", text) and not re.fullmatch(r"[\W_]+", text):
        print("  ✅ Contains non-ASCII characters")
        return True
    
    # For English text with spaces (potential phrases/sentences)
    if " " in text and re.search(r"[a-zA-Z]", text):
        print("  ✓ Contains spaces and letters")
        # Skip simple labels like "Item 1", "Page 2"
        if re.fullmatch(r"[A-Za-z]+\s*\d+|\d+\s*[A-Za-z]+", text):
            print("  ❌ Simple labels")
            return False
        # Skip short combinations like "ID ABC123"
        if len(text.split()) <= 2 and re.search(r"[A-Z0-9]+", text):
            print("  ❌ Short combinations")
            return False
        # Translate longer English phrases (3+ words or complex content)
        if len(text.split()) >= 3 or len(text) > 20:
            print("  ✅ Long phrases or complex content")
            return True
    
    # Default: don't translate
    print("  ❌ Default: don't translate")
    return False

def test_debug():
    """Test the debug function"""
    test_words = ["Name", "Country", "Language", "Alice", "USA", "English", "Bob", "France", "French"]
    
    for word in test_words:
        result = debug_should_translate_text(word)
        print(f"Final result for '{word}': {result}")
        print("-" * 50)

if __name__ == "__main__":
    test_debug()