"""
Exception classes for Offitrans

This module defines custom exceptions used throughout the Offitrans library.
"""

from .errors import (
    OffitransError,
    TranslationError,
    ProcessorError,
    ConfigError,
    FileError,
    APIError,
    CacheError
)

__all__ = [
    'OffitransError',
    'TranslationError', 
    'ProcessorError',
    'ConfigError',
    'FileError',
    'APIError',
    'CacheError',
]