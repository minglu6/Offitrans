"""
Utility functions for Offitrans

This module provides common utility functions for text processing,
language detection, and validation.
"""

import re
import logging
from typing import List, Tuple, Dict, Optional, Set
from pathlib import Path

logger = logging.getLogger(__name__)


def detect_language(text: str) -> str:
    """
    Detect the language of the given text.

    Args:
        text: Text to analyze for language detection

    Returns:
        Language code (e.g., 'zh', 'en', 'th', etc.) or 'unknown'
    """
    if not text or not text.strip():
        return "unknown"

    text = text.strip()

    # Chinese characters
    if re.search(r"[\u4e00-\u9fff]", text):
        return "zh"

    # Thai characters
    if re.search(r"[\u0e00-\u0e7f]", text):
        return "th"

    # Japanese characters (Hiragana, Katakana, Kanji)
    if re.search(r"[\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf\uff66-\uff9f]", text):
        return "ja"

    # Korean characters
    if re.search(
        r"[\uac00-\ud7af\u1100-\u11ff\u3130-\u318f\ua960-\ua97f\ud7b0-\ud7ff]", text
    ):
        return "ko"

    # Arabic characters
    if re.search(r"[\u0600-\u06ff\u0750-\u077f]", text):
        return "ar"

    # Russian/Cyrillic characters
    if re.search(r"[\u0400-\u04ff]", text):
        return "ru"

    # German specific characters
    if re.search(r"[äöüßÄÖÜ]", text):
        return "de"

    # French specific characters
    if re.search(r"[àâäéèêëïîôùûüÿçÀÂÄÉÈÊËÏÎÔÙÛÜŸÇ]", text):
        return "fr"

    # Spanish specific characters
    if re.search(r"[ñáéíóúüÑÁÉÍÓÚÜ¿¡]", text):
        return "es"

    # If contains mainly Latin characters, assume English
    if re.search(r"[a-zA-Z]", text):
        return "en"

    return "unknown"


def validate_language_code(
    lang_code: str, supported_languages: Optional[Dict[str, str]] = None
) -> bool:
    """
    Validate if a language code is supported.

    Args:
        lang_code: Language code to validate
        supported_languages: Dictionary of supported language codes

    Returns:
        True if language code is valid, False otherwise
    """
    if not lang_code:
        return False

    # Default supported languages if not provided
    if supported_languages is None:
        supported_languages = {
            "zh": "Chinese",
            "en": "English",
            "th": "Thai",
            "ja": "Japanese",
            "ko": "Korean",
            "fr": "French",
            "de": "German",
            "es": "Spanish",
            "ar": "Arabic",
            "ru": "Russian",
            "auto": "Auto-detect",
        }

    return lang_code.lower() in supported_languages


def clean_text(text: str) -> str:
    """
    Clean and normalize text for translation.

    Args:
        text: Raw text to clean

    Returns:
        Cleaned and normalized text
    """
    if not text:
        return text

    # Remove excessive whitespace
    cleaned = re.sub(r"\s+", " ", text.strip())

    # Remove control characters but keep line breaks and tabs
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", cleaned)

    # Normalize quotes
    cleaned = re.sub(r'[""' "`]", '"', cleaned)
    cleaned = re.sub(r"[" "`]", "'", cleaned)

    return cleaned


def should_translate_text(text: str) -> bool:
    """
    Determine if a text should be translated based on content analysis.

    Args:
        text: Text to analyze

    Returns:
        True if text should be translated, False otherwise
    """
    if not text or not text.strip():
        return False

    text = text.strip()

    # Skip pure numbers
    if text.isdigit():
        return False

    # Skip pure symbols
    if re.fullmatch(r"[\W_]+", text):
        return False

    # Skip pure English letters (single words without spaces)
    if re.fullmatch(r"[a-zA-Z]+", text) and len(text) <= 10:
        return False

    # Skip alphanumeric combinations (like IDs, codes)
    if re.fullmatch(r"[a-zA-Z0-9]+", text):
        return False

    # Skip numbers with symbols (prices, percentages, measurements)
    if re.fullmatch(r"[\d\W_]+", text):
        return False

    # Skip URLs and emails
    if re.search(r"https?://|www\.|@.*\.|\.com|\.org|\.net|\.edu", text.lower()):
        return False

    # Skip file paths
    if re.search(
        r"[A-Za-z]:\\|/[a-zA-Z]|\.exe|\.dll|\.pdf|\.docx?|\.xlsx?|\.pptx?", text
    ):
        return False

    # Skip programming identifiers (underscore or camelCase)
    if re.search(r"[a-zA-Z]+_[a-zA-Z]+|[a-z]+[A-Z][a-z]*", text):
        return False

    # Skip measurements and units
    if re.fullmatch(
        r"\d+\s*(mm|cm|m|km|kg|g|ml|l|°C|°F|%|px|pt|em|rem|in|ft)", text, re.IGNORECASE
    ):
        return False

    # Skip version numbers
    if re.search(r"v\d+\.\d+|ver\.\d+|version\s*\d+", text.lower()):
        return False

    # Skip date formats
    if re.search(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4}", text):
        return False

    # Skip time formats
    if re.search(r"\d{1,2}:\d{2}(\s*(AM|PM))?", text.upper()):
        return False

    # Skip formulas (starting with =)
    if text.startswith("="):
        return False

    # Translate if contains Chinese characters
    if re.search(r"[\u4e00-\u9fff]", text):
        return True

    # Translate if contains other non-ASCII characters (except symbols)
    if re.search(r"[^\x00-\x7f]", text) and not re.fullmatch(r"[\W_]+", text):
        return True

    # For English text with spaces (potential phrases/sentences)
    if " " in text and re.search(r"[a-zA-Z]", text):
        # Skip simple labels like "Item 1", "Page 2"
        if re.fullmatch(r"[A-Za-z]+\s*\d+|\d+\s*[A-Za-z]+", text):
            return False
        # Skip short combinations like "ID ABC123"
        if len(text.split()) <= 2 and re.search(r"[A-Z0-9]+", text):
            return False
        # Translate longer English phrases (3+ words or complex content)
        if len(text.split()) >= 3 or len(text) > 20:
            return True

    # Default: don't translate
    return False


def split_text_chunks(
    text: str, max_chunk_size: int = 5000, overlap: int = 100
) -> List[str]:
    """
    Split large text into smaller chunks for translation.

    Args:
        text: Text to split
        max_chunk_size: Maximum size of each chunk in characters
        overlap: Number of characters to overlap between chunks

    Returns:
        List of text chunks
    """
    if not text or len(text) <= max_chunk_size:
        return [text] if text else []

    chunks = []
    start = 0

    while start < len(text):
        # Find the end of this chunk
        end = min(start + max_chunk_size, len(text))

        # Try to break at sentence boundaries
        if end < len(text):
            # Look for sentence endings within the last 200 characters
            sentence_endings = [".", "!", "?", "。", "！", "？"]
            for i in range(end - 200, end):
                if i > start and text[i] in sentence_endings:
                    # Check if it's really a sentence ending (not abbreviation)
                    if i + 1 < len(text) and text[i + 1] in [" ", "\n", "\t"]:
                        end = i + 1
                        break

            # If no sentence ending found, try to break at word boundaries
            else:
                for i in range(end - 50, end):
                    if i > start and text[i] in [" ", "\n", "\t"]:
                        end = i
                        break

        chunk = text[start:end].strip()
        if chunk:
            chunks.append(chunk)

        # Move start position (with overlap if not at the end)
        if end >= len(text):
            break
        start = max(start + 1, end - overlap)

    return chunks


def filter_translatable_texts(texts: List[str]) -> Tuple[List[str], List[str]]:
    """
    Filter texts into translatable and non-translatable lists.

    Args:
        texts: List of texts to filter

    Returns:
        Tuple of (translatable_texts, non_translatable_texts)
    """
    translatable = []
    non_translatable = []

    for text in texts:
        if should_translate_text(text):
            translatable.append(text)
        else:
            non_translatable.append(text)

    return translatable, non_translatable


def deduplicate_texts(texts: List[str]) -> Tuple[List[str], Dict[str, List[int]]]:
    """
    Remove duplicate texts and return mapping of unique texts to original indices.

    Args:
        texts: List of texts that may contain duplicates

    Returns:
        Tuple of (unique_texts, text_to_indices_mapping)
    """
    seen_texts: Set[str] = set()
    unique_texts: List[str] = []
    text_to_indices: Dict[str, List[int]] = {}

    for i, text in enumerate(texts):
        if text not in seen_texts:
            seen_texts.add(text)
            unique_texts.append(text)
            text_to_indices[text] = [i]
        else:
            text_to_indices[text].append(i)

    return unique_texts, text_to_indices


def normalize_text(text: str) -> str:
    """
    Normalize text for better deduplication and comparison.

    Args:
        text: Raw text to normalize

    Returns:
        Normalized text
    """
    if not text:
        return text

    # Remove extra whitespace and normalize
    normalized = re.sub(r"\s+", " ", text.strip())

    # Convert to lowercase for comparison (but keep original case)
    return normalized


def get_file_encoding(file_path: str) -> str:
    """
    Detect file encoding.

    Args:
        file_path: Path to the file

    Returns:
        Detected encoding or 'utf-8' as default
    """
    try:
        import chardet

        with open(file_path, "rb") as f:
            raw_data = f.read(10000)  # Read first 10KB
            result = chardet.detect(raw_data)
            return result.get("encoding", "utf-8")

    except ImportError:
        logger.warning("chardet not available, using utf-8 encoding")
        return "utf-8"
    except Exception as e:
        logger.error(f"Error detecting encoding for {file_path}: {e}")
        return "utf-8"


def safe_filename(filename: str) -> str:
    """
    Create a safe filename by removing/replacing invalid characters.

    Args:
        filename: Original filename

    Returns:
        Safe filename for filesystem use
    """
    # Replace invalid characters with underscores
    safe_name = re.sub(r'[<>:"/\\|?*]', "_", filename)

    # Remove control characters
    safe_name = re.sub(r"[\x00-\x1f\x7f]", "", safe_name)

    # Limit length and strip spaces/dots from ends
    safe_name = safe_name[:255].strip(" .")

    # Ensure it's not empty
    if not safe_name:
        safe_name = "unnamed_file"

    return safe_name


# Backward compatibility aliases
should_translate = should_translate_text
filter_texts = filter_translatable_texts
