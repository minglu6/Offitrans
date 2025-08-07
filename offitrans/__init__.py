"""
Offitrans - Office File Translation Library

A powerful Office file translation tool library that supports batch translation
of PDF, Excel, PPT, and Word documents.

Author: Offitrans Contributors
License: MIT
"""

from .version import __version__

# Core components
from .core.base import BaseTranslator
from .core.cache import TranslationCache
from .core.config import Config

# Main translators
from .translators.google import GoogleTranslator

# File processors
from .processors.excel import ExcelProcessor
from .processors.word import WordProcessor
from .processors.pdf import PDFProcessor
from .processors.powerpoint import PowerPointProcessor

# Exceptions
from .exceptions.errors import (
    OffitransError,
    TranslationError,
    ProcessorError,
    ConfigError,
)

__all__ = [
    # Version
    "__version__",
    # Core
    "BaseTranslator",
    "TranslationCache",
    "Config",
    # Translators
    "GoogleTranslator",
    # Processors
    "ExcelProcessor",
    "WordProcessor",
    "PDFProcessor",
    "PowerPointProcessor",
    # Exceptions
    "OffitransError",
    "TranslationError",
    "ProcessorError",
    "ConfigError",
]

# Backward compatibility aliases
ExcelTranslator = ExcelProcessor  # Keep old name for compatibility
