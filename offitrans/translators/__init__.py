"""
Translation engines for Offitrans

This module contains various translation service implementations.
"""

from .google import GoogleTranslator
from .base_api import BaseAPITranslator

__all__ = [
    "GoogleTranslator",
    "BaseAPITranslator",
]

# Available translator types
AVAILABLE_TRANSLATORS = {
    "google": GoogleTranslator,
}


def get_translator(translator_type: str, **kwargs):
    """
    Get a translator instance by type.

    Args:
        translator_type: Type of translator ('google', etc.)
        **kwargs: Arguments to pass to translator constructor

    Returns:
        Translator instance

    Raises:
        ValueError: If translator type is not available
    """
    if translator_type not in AVAILABLE_TRANSLATORS:
        available = ", ".join(AVAILABLE_TRANSLATORS.keys())
        raise ValueError(
            f"Unknown translator type: {translator_type}. Available: {available}"
        )

    translator_class = AVAILABLE_TRANSLATORS[translator_type]
    return translator_class(**kwargs)
