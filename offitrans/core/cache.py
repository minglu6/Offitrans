"""
Translation cache management for Offitrans

This module provides caching functionality to reduce redundant translation API calls
and improve performance by storing previously translated text.
"""

import json
import os
import hashlib
import atexit
import threading
import logging
from typing import Dict, Optional, Any, List
from functools import wraps
from pathlib import Path

logger = logging.getLogger(__name__)


class TranslationCache:
    """
    Translation cache manager that stores and retrieves translated text
    to reduce API calls and improve performance.
    """
    
    def __init__(self, cache_file: Optional[str] = None, auto_save_interval: int = 10):
        """
        Initialize translation cache.
        
        Args:
            cache_file: Path to the cache file (default: uses XDG cache directory)
            auto_save_interval: Save cache every N operations (default: 10)
        """
        if cache_file is None:
            from .config import get_default_cache_path
            cache_file = get_default_cache_path()
        self.cache_file = Path(cache_file)
        self.auto_save_interval = auto_save_interval
        self._cache: Dict[str, str] = {}
        self._lock = threading.Lock()
        self._operation_count = 0
        
        # Create cache directory if it doesn't exist
        self.cache_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Load existing cache
        self._load_cache()
        
        # Register exit handler to save cache on program termination
        atexit.register(self._save_cache_on_exit)
    
    def _generate_key(self, text: str, source_lang: str, target_lang: str) -> str:
        """
        Generate a unique cache key for the translation.
        
        Args:
            text: Original text
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            MD5 hash as cache key
        """
        # Normalize the key string
        key_string = f"{source_lang.lower()}:{target_lang.lower()}:{text.strip()}"
        return hashlib.md5(key_string.encode('utf-8')).hexdigest()
    
    def _load_cache(self) -> None:
        """Load cache from file."""
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    self._cache = json.load(f)
                logger.info(f"Translation cache loaded: {len(self._cache)} entries")
            else:
                self._cache = {}
                logger.info("Created new translation cache")
        except Exception as e:
            logger.error(f"Failed to load cache file: {e}")
            self._cache = {}
    
    def _save_cache(self) -> None:
        """Save cache to file."""
        try:
            with self._lock:
                # Create a backup if cache file exists
                if self.cache_file.exists():
                    backup_file = self.cache_file.with_suffix('.bak')
                    self.cache_file.rename(backup_file)
                
                # Write new cache file
                with open(self.cache_file, 'w', encoding='utf-8') as f:
                    json.dump(self._cache, f, ensure_ascii=False, indent=2)
                
                # Remove backup on successful write
                backup_file = self.cache_file.with_suffix('.bak')
                if backup_file.exists():
                    backup_file.unlink()
                    
                logger.debug(f"Cache saved: {len(self._cache)} entries")
        except Exception as e:
            logger.error(f"Failed to save cache file: {e}")
            # Restore backup if it exists
            backup_file = self.cache_file.with_suffix('.bak')
            if backup_file.exists():
                backup_file.rename(self.cache_file)
    
    def _save_cache_on_exit(self) -> None:
        """Save cache when program exits."""
        try:
            self._save_cache()
            logger.info(f"Translation cache saved on exit: {len(self._cache)} entries")
        except Exception as e:
            logger.error(f"Failed to save cache on exit: {e}")
    
    def get(self, text: str, source_lang: str, target_lang: str) -> Optional[str]:
        """
        Get translation from cache.
        
        Args:
            text: Original text
            source_lang: Source language code  
            target_lang: Target language code
            
        Returns:
            Cached translation if exists, None otherwise
        """
        if not text or not text.strip():
            return text
            
        key = self._generate_key(text, source_lang, target_lang)
        with self._lock:
            return self._cache.get(key)
    
    def set(self, text: str, translation: str, source_lang: str, target_lang: str, 
            force_save: bool = False) -> None:
        """
        Set translation in cache.
        
        Args:
            text: Original text
            translation: Translated text
            source_lang: Source language code
            target_lang: Target language code
            force_save: Force immediate save to disk
        """
        if not text or not text.strip() or not translation:
            return
            
        key = self._generate_key(text, source_lang, target_lang)
        with self._lock:
            self._cache[key] = translation
            self._operation_count += 1
        
        # Auto-save based on interval or force_save flag
        if force_save or self._operation_count >= self.auto_save_interval:
            self._save_cache()
            self._operation_count = 0
    
    def clear(self) -> None:
        """Clear all cached translations."""
        with self._lock:
            self._cache.clear()
            self._operation_count = 0
        self._save_cache()
        logger.info("Translation cache cleared")
    
    def get_batch(self, texts: List[str], source_lang: str, target_lang: str) -> Dict[str, Optional[str]]:
        """
        Get multiple translations from cache.
        
        Args:
            texts: List of original texts
            source_lang: Source language code
            target_lang: Target language code
            
        Returns:
            Dictionary mapping original text to cached translation (None if not cached)
        """
        result = {}
        for text in texts:
            if text and text.strip():
                cached_result = self.get(text, source_lang, target_lang)
                result[text] = cached_result
            else:
                result[text] = text  # Return empty text as-is
        return result
    
    def set_batch(self, text_translation_pairs: Dict[str, str], 
                  source_lang: str, target_lang: str) -> None:
        """
        Set multiple translations in cache and save immediately.
        
        Args:
            text_translation_pairs: Dictionary of original text to translation
            source_lang: Source language code  
            target_lang: Target language code
        """
        count = 0
        for text, translation in text_translation_pairs.items():
            if text and text.strip() and translation and translation != text:
                key = self._generate_key(text, source_lang, target_lang)
                with self._lock:
                    self._cache[key] = translation
                count += 1
        
        # Save immediately after batch operation
        if count > 0:
            self._save_cache() 
            self._operation_count = 0
            logger.info(f"Batch cache saved: {count} entries")

    def save(self) -> None:
        """Manually save cache to disk."""
        self._save_cache()
        logger.info(f"Translation cache manually saved: {len(self._cache)} entries")
    
    def get_stats(self) -> Dict[str, Any]:
        """
        Get cache statistics.
        
        Returns:
            Dictionary containing cache statistics
        """
        with self._lock:
            cache_size = len(self._cache)
            
        file_size = 0
        if self.cache_file.exists():
            file_size = self.cache_file.stat().st_size
            
        return {
            "total_entries": cache_size,
            "cache_file": str(self.cache_file),
            "file_exists": self.cache_file.exists(),
            "file_size_bytes": file_size,
            "pending_operations": self._operation_count,
            "auto_save_interval": self.auto_save_interval
        }
    
    def cleanup_old_entries(self, max_entries: int = 10000) -> int:
        """
        Remove old cache entries if cache grows too large.
        
        Args:
            max_entries: Maximum number of entries to keep
            
        Returns:
            Number of entries removed
        """
        with self._lock:
            if len(self._cache) <= max_entries:
                return 0
                
            # Keep the most recently accessed entries (simple FIFO for now)
            entries_to_remove = len(self._cache) - max_entries
            keys_to_remove = list(self._cache.keys())[:entries_to_remove]
            
            for key in keys_to_remove:
                del self._cache[key]
        
        self._save_cache()
        logger.info(f"Cleaned up {entries_to_remove} old cache entries")
        return entries_to_remove

    def __len__(self) -> int:
        """Return number of cached entries."""
        with self._lock:
            return len(self._cache)
    
    def __contains__(self, key_tuple) -> bool:
        """Check if a translation is cached."""
        if isinstance(key_tuple, tuple) and len(key_tuple) == 3:
            text, source_lang, target_lang = key_tuple
            return self.get(text, source_lang, target_lang) is not None
        return False


# Global cache instance
_global_cache = TranslationCache()


def cached_translation(cache_instance: Optional[TranslationCache] = None):
    """
    Decorator for caching translation results.
    
    Args:
        cache_instance: Cache instance to use (uses global cache if None)
        
    Returns:
        Decorated translation function with caching
    """
    def decorator(translate_func):
        @wraps(translate_func)
        def wrapper(self, text: str) -> str:
            # Check if caching is enabled for this translator
            if not getattr(self, 'enable_cache', True):
                logger.debug(f"Cache disabled, calling API directly: {text[:50]}...")
                return translate_func(self, text)
            
            # Use specified cache instance or global cache
            cache = cache_instance or _global_cache
            
            # Try to get from cache first
            cached_result = cache.get(text, self.source_lang, self.target_lang)
            if cached_result is not None:
                logger.debug(f"Cache hit: {text[:50]}...")
                return cached_result
            
            # Cache miss - call original translation function
            logger.debug(f"Cache miss, calling API: {text[:50]}...")
            result = translate_func(self, text)
            
            # Store result in cache (force save for individual translations)
            if result and result != text:  # Only cache successful translations
                cache.set(text, result, self.source_lang, self.target_lang, force_save=True)
            
            return result
        return wrapper
    return decorator


def get_global_cache() -> TranslationCache:
    """
    Get the global cache instance.
    
    Returns:
        Global TranslationCache instance
    """
    return _global_cache


def set_global_cache_file(cache_file: str) -> None:
    """
    Set the global cache file path.
    
    Args:
        cache_file: Path to the cache file
    """
    global _global_cache
    _global_cache = TranslationCache(cache_file)