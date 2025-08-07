"""
Base translator class for Offitrans

This module provides the abstract base class for all translators.
"""

from abc import ABC, abstractmethod
from typing import Optional, List, Dict, Any
import threading
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

from ..exceptions.errors import TranslationError

logger = logging.getLogger(__name__)


class BaseTranslator(ABC):
    """
    Abstract base class for all translators.
    
    This class defines the common interface and functionality that all
    translator implementations must provide.
    """

    def __init__(self, 
                 source_lang: str = "auto",
                 target_lang: str = "en",
                 max_workers: int = 5,
                 timeout: int = 120,
                 retry_count: int = 3,
                 retry_delay: int = 2,
                 batch_size: int = 20,
                 enable_cache: bool = True,
                 **kwargs):
        """
        Initialize the base translator.
        
        Args:
            source_lang: Source language code (default: "auto" for auto-detection)
            target_lang: Target language code (default: "en")  
            max_workers: Maximum number of concurrent workers (default: 5)
            timeout: Request timeout in seconds (default: 120)
            retry_count: Number of retry attempts (default: 3)
            retry_delay: Delay between retries in seconds (default: 2)
            batch_size: Batch processing size (default: 20)
            enable_cache: Whether to enable translation cache (default: True)
            **kwargs: Additional keyword arguments
        """
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.max_workers = max_workers
        self.timeout = timeout
        self.retry_count = retry_count
        self.retry_delay = retry_delay
        self.batch_size = batch_size
        self.enable_cache = enable_cache
        
        # Supported languages mapping
        self.supported_languages = {
            'zh': 'zh',  # Chinese
            'en': 'en',  # English
            'th': 'th',  # Thai
            'ja': 'ja',  # Japanese
            'ko': 'ko',  # Korean
            'fr': 'fr',  # French
            'de': 'de',  # German
            'es': 'es',  # Spanish
            'auto': 'auto',  # Auto-detection
        }
        
        # Thread safety
        self._lock = threading.Lock()
        
        # Statistics
        self.stats = {
            'total_translations': 0,
            'successful_translations': 0,
            'failed_translations': 0,
            'total_chars_translated': 0,
        }
        
        # Initialize any additional settings from kwargs
        self._init_kwargs(kwargs)

    def _init_kwargs(self, kwargs: Dict[str, Any]) -> None:
        """
        Initialize additional settings from keyword arguments.
        
        Args:
            kwargs: Additional keyword arguments
        """
        # Override supported languages if provided
        if 'supported_languages' in kwargs:
            self.supported_languages.update(kwargs['supported_languages'])
        
        # Set additional configuration
        for key, value in kwargs.items():
            if not hasattr(self, key):
                setattr(self, key, value)

    @abstractmethod
    def translate_text(self, text: str) -> str:
        """
        Translate a single text string from source language to target language.

        Args:
            text: The text to translate.

        Returns:
            Translated text.
            
        Raises:
            TranslationError: If translation fails
        """
        pass

    def translate_text_batch(self, texts: List[str]) -> List[str]:
        """
        Translate a batch of text strings from source language to target language.
        Uses multithreading for improved performance.

        Args:
            texts: List of text strings to translate.

        Returns:
            List of translated text strings.
            
        Raises:
            TranslationError: If batch translation fails
        """
        if not texts:
            return []

        logger.info(f"Starting batch translation of {len(texts)} texts")
        
        # Use multithreading for translation
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all translation tasks
            future_to_index = {
                executor.submit(self.translate_text, text): i 
                for i, text in enumerate(texts)
            }

            # Initialize results list with proper typing
            results: List[str] = [""] * len(texts)

            # Collect results
            for future in as_completed(future_to_index):
                index = future_to_index[future]
                try:
                    result = future.result()
                    results[index] = result if result is not None else ""
                    self._update_stats(success=True, chars=len(texts[index]))
                except Exception as exc:
                    logger.error(f'Translation at index {index} failed: {exc}')
                    results[index] = texts[index]  # Return original text on error
                    self._update_stats(success=False)

            logger.info(f"Batch translation completed: {len(results)} results")
            return results

    def translate_text_batch_simple(self, texts: List[str]) -> List[str]:
        """
        Simple multithreaded version using map (for backward compatibility).

        Args:
            texts: List of text strings to translate.

        Returns:
            List of translated text strings.
        """
        if not texts:
            return []

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            return list(executor.map(self.translate_text, texts))

    def validate_language_code(self, lang_code: str) -> bool:
        """
        Validate if a language code is supported.
        
        Args:
            lang_code: Language code to validate
            
        Returns:
            True if supported, False otherwise
        """
        return lang_code in self.supported_languages

    def get_supported_languages(self) -> Dict[str, str]:
        """
        Get the dictionary of supported languages.
        
        Returns:
            Dictionary mapping language codes to language names
        """
        return self.supported_languages.copy()

    def _update_stats(self, success: bool = True, chars: int = 0) -> None:
        """
        Update translation statistics (thread-safe).
        
        Args:
            success: Whether the translation was successful
            chars: Number of characters translated
        """
        with self._lock:
            self.stats['total_translations'] += 1
            if success:
                self.stats['successful_translations'] += 1
                self.stats['total_chars_translated'] += chars
            else:
                self.stats['failed_translations'] += 1

    def get_stats(self) -> Dict[str, int]:
        """
        Get translation statistics.
        
        Returns:
            Dictionary containing translation statistics
        """
        with self._lock:
            return self.stats.copy()

    def reset_stats(self) -> None:
        """Reset translation statistics."""
        with self._lock:
            self.stats = {
                'total_translations': 0,
                'successful_translations': 0,
                'failed_translations': 0,
                'total_chars_translated': 0,
            }

    def __str__(self) -> str:
        """String representation of the translator."""
        return f"{self.__class__.__name__}({self.source_lang} -> {self.target_lang})"

    def __repr__(self) -> str:
        """Detailed string representation of the translator."""
        return (f"{self.__class__.__name__}("
                f"source_lang='{self.source_lang}', "
                f"target_lang='{self.target_lang}', "
                f"max_workers={self.max_workers})")


# Backward compatibility alias
Translator = BaseTranslator