"""
Unit tests for core functionality
"""

import pytest
from unittest.mock import Mock, patch

from offitrans.core.base import BaseTranslator
from offitrans.core.cache import TranslationCache, get_global_cache
from offitrans.core.config import Config
from offitrans.core.utils import (
    detect_language,
    validate_language_code,
    clean_text,
    should_translate_text,
    split_text_chunks,
    filter_translatable_texts,
    deduplicate_texts
)
from offitrans.exceptions.errors import TranslationError, ConfigError


class TestTranslator(BaseTranslator):
    """Concrete implementation of BaseTranslator for testing"""
    
    def translate_text(self, text: str) -> str:
        return f"translated_{text}"


class TestBaseTranslator:
    """Test the BaseTranslator class"""
    
    def test_initialization(self):
        """Test translator initialization"""
        translator = TestTranslator()
        assert translator.source_lang == "auto"
        assert translator.target_lang == "en"
        assert translator.max_workers == 5
        assert translator.timeout == 120
    
    def test_custom_initialization(self):
        """Test translator with custom parameters"""
        translator = TestTranslator(
            source_lang="zh",
            target_lang="th",
            max_workers=3,
            timeout=60
        )
        assert translator.source_lang == "zh"
        assert translator.target_lang == "th"
        assert translator.max_workers == 3
        assert translator.timeout == 60
    
    def test_validate_language_code(self):
        """Test language code validation"""
        translator = TestTranslator()
        assert translator.validate_language_code("en") is True
        assert translator.validate_language_code("zh") is True
        assert translator.validate_language_code("invalid") is False
    
    def test_translate_text_batch(self):
        """Test batch translation"""
        translator = TestTranslator()
        texts = ["hello", "world", "test"]
        result = translator.translate_text_batch(texts)
        
        assert len(result) == 3
        assert all("translated_" in text for text in result)
    
    def test_translate_empty_batch(self):
        """Test batch translation with empty list"""
        translator = TestTranslator()
        result = translator.translate_text_batch([])
        assert result == []
    
    def test_get_stats(self):
        """Test statistics tracking"""
        translator = TestTranslator()
        stats = translator.get_stats()
        
        assert "total_translations" in stats
        assert "successful_translations" in stats
        assert "failed_translations" in stats
        assert stats["total_translations"] == 0
    
    def test_reset_stats(self):
        """Test statistics reset"""
        translator = TestTranslator()
        # Trigger some translations to update stats
        translator.translate_text_batch(["test"])
        
        translator.reset_stats()
        stats = translator.get_stats()
        assert stats["total_translations"] == 0


class TestTranslationCache:
    """Test the TranslationCache class"""
    
    def test_cache_initialization(self, temp_dir):
        """Test cache initialization"""
        cache_file = temp_dir / "test_cache.json"
        cache = TranslationCache(str(cache_file))
        
        assert cache.cache_file.name == "test_cache.json"
        assert len(cache) == 0
    
    def test_cache_set_get(self, temp_dir):
        """Test setting and getting cache entries"""
        cache_file = temp_dir / "test_cache.json"
        cache = TranslationCache(str(cache_file))
        
        # Set a translation
        cache.set("hello", "hola", "en", "es")
        
        # Get the translation
        result = cache.get("hello", "en", "es")
        assert result == "hola"
    
    def test_cache_miss(self, temp_dir):
        """Test cache miss"""
        cache_file = temp_dir / "test_cache.json"
        cache = TranslationCache(str(cache_file))
        
        result = cache.get("nonexistent", "en", "zh")
        assert result is None
    
    def test_cache_batch_operations(self, temp_dir):
        """Test batch cache operations"""
        cache_file = temp_dir / "test_cache.json"
        cache = TranslationCache(str(cache_file))
        
        # Set batch
        translations = {
            "hello": "hola",
            "world": "mundo",
            "test": "prueba"
        }
        cache.set_batch(translations, "en", "es")
        
        # Get batch
        texts = ["hello", "world", "test", "missing"]
        results = cache.get_batch(texts, "en", "es")
        
        assert results["hello"] == "hola"
        assert results["world"] == "mundo"
        assert results["test"] == "prueba"
        assert results["missing"] is None
    
    def test_cache_clear(self, temp_dir):
        """Test cache clearing"""
        cache_file = temp_dir / "test_cache.json"
        cache = TranslationCache(str(cache_file))
        
        cache.set("hello", "hola", "en", "es")
        assert len(cache) == 1
        
        cache.clear()
        assert len(cache) == 0
    
    def test_cache_stats(self, temp_dir):
        """Test cache statistics"""
        cache_file = temp_dir / "test_cache.json"
        cache = TranslationCache(str(cache_file))
        
        cache.set("hello", "hola", "en", "es")
        stats = cache.get_stats()
        
        assert stats["total_entries"] == 1
        assert stats["cache_file"] == str(cache_file)
        assert "file_exists" in stats


class TestConfig:
    """Test the Config class"""
    
    def test_default_config(self):
        """Test default configuration"""
        config = Config()
        
        assert config.translator.max_workers == 5
        assert config.translator.timeout == 120
        assert config.cache.enabled is True
        assert config.processor.preserve_formatting is True
    
    def test_config_update(self):
        """Test configuration updates"""
        config = Config()
        
        config.update(max_workers=10, timeout=60)
        
        assert config.translator.max_workers == 10
        assert config.translator.timeout == 60
    
    def test_config_validation(self):
        """Test configuration validation"""
        config = Config()
        
        # Valid configuration
        assert config.validate() is True
        
        # Invalid configuration
        config.translator.max_workers = -1
        assert config.validate() is False
    
    def test_config_save_load(self, temp_dir):
        """Test configuration save and load"""
        config_file = temp_dir / "test_config.json"
        
        # Create and save config
        config1 = Config()
        config1.translator.max_workers = 10
        config1.save_to_file(str(config_file))
        
        # Load config
        config2 = Config(str(config_file))
        assert config2.translator.max_workers == 10


class TestUtilityFunctions:
    """Test utility functions"""
    
    def test_detect_language(self):
        """Test language detection"""
        assert detect_language("Hello world") == "en"
        assert detect_language("Hola mundo") == "en"  # No special Spanish characters, defaults to English
        assert detect_language("¡Hola! ¿Cómo estás?") == "es"  # With Spanish punctuation
        assert detect_language("สวัสดี") == "th"
        assert detect_language("123") == "unknown"
        assert detect_language("") == "unknown"
    
    def test_validate_language_code(self):
        """Test language code validation"""
        assert validate_language_code("en") is True
        assert validate_language_code("zh") is True
        assert validate_language_code("invalid") is False
        assert validate_language_code("") is False
    
    def test_clean_text(self):
        """Test text cleaning"""
        assert clean_text("  hello   world  ") == "hello world"
        assert clean_text("hello\x00world") == "helloworld"
        assert clean_text("hello\n\nworld") == "hello world"
    
    def test_should_translate_text(self):
        """Test translation necessity detection"""
        assert should_translate_text("Hello world") is False  # Pure English, skipped by design
        assert should_translate_text("¡Hola! ¿Cómo estás?") is True  # Spanish with punctuation
        assert should_translate_text("123") is False
        assert should_translate_text("test@email.com") is False
        assert should_translate_text("") is False
    
    def test_split_text_chunks(self):
        """Test text chunking"""
        long_text = "This is a very long text. " * 100
        chunks = split_text_chunks(long_text, max_chunk_size=100)
        
        assert len(chunks) > 1
        assert all(len(chunk) <= 150 for chunk in chunks)  # Allow some overlap
    
    def test_filter_translatable_texts(self):
        """Test text filtering"""
        texts = [
            "Hello world",      # Not translatable (pure English)
            "¡Hola! ¿Cómo estás?",  # Translatable (Spanish with punctuation)
            "123",              # Not translatable
            "test@email.com"    # Not translatable
        ]
        
        translatable, non_translatable = filter_translatable_texts(texts)
        
        assert len(translatable) == 1
        assert len(non_translatable) == 3
        assert "¡Hola! ¿Cómo estás?" in translatable
        assert "Hello world" in non_translatable
        assert "123" in non_translatable
        assert "test@email.com" in non_translatable
    
    def test_deduplicate_texts(self):
        """Test text deduplication"""
        texts = ["hello", "world", "hello", "test", "world"]
        unique_texts, mapping = deduplicate_texts(texts)
        
        assert len(unique_texts) == 3
        assert set(unique_texts) == {"hello", "world", "test"}
        assert len(mapping["hello"]) == 2  # Appears at indices 0 and 2
        assert len(mapping["world"]) == 2  # Appears at indices 1 and 4


def test_global_cache():
    """Test global cache functionality"""
    global_cache = get_global_cache()
    assert global_cache is not None
    
    # Test that it's a singleton
    global_cache2 = get_global_cache()
    assert global_cache is global_cache2