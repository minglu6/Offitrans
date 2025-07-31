"""
Core module for Offitrans

This module contains the base classes and core functionality for the translation system.
"""

from .base import BaseTranslator
from .cache import TranslationCache, cached_translation
from .config import Config
from .utils import (
    detect_language,
    validate_language_code,
    clean_text,
    split_text_chunks
)

__all__ = [
    'BaseTranslator',
    'TranslationCache',
    'cached_translation',
    'Config',
    'detect_language',
    'validate_language_code', 
    'clean_text',
    'split_text_chunks',
]