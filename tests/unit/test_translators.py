"""
Unit tests for translator classes
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
import requests

from offitrans.translators.google import GoogleTranslator, get_supported_languages
from offitrans.translators.base_api import BaseAPITranslator
from offitrans.exceptions.errors import TranslationError, ConfigError


class TestGoogleTranslator:
    """Test the GoogleTranslator class"""

    def test_initialization_free_api(self):
        """Test Google translator initialization with free API"""
        translator = GoogleTranslator(use_free_api=True)

        assert translator.use_free_api is True
        assert "translate.googleapis.com" in translator.api_url
        assert translator.source_lang == "auto"
        assert translator.target_lang == "en"

    def test_initialization_paid_api(self):
        """Test Google translator initialization with paid API"""
        translator = GoogleTranslator(api_key="test_key", use_free_api=False)

        assert translator.use_free_api is False
        assert translator.api_key == "test_key"
        assert "translation.googleapis.com" in translator.api_url

    @patch("requests.get")
    def test_translate_free_api_success(self, mock_get):
        """Test successful translation with free API"""
        # Mock successful response
        mock_response = Mock()
        mock_response.json.return_value = [
            [["Hello", "Hola", None, None, None, None, None, None, []]]
        ]
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response

        translator = GoogleTranslator(use_free_api=True)
        result = translator._translate_free_api("Hola")

        assert result == "Hello"
        mock_get.assert_called_once()

    @patch("requests.get")
    def test_translate_free_api_failure(self, mock_get):
        """Test translation failure with free API"""
        # Mock failed response
        mock_get.side_effect = requests.exceptions.RequestException("Network error")

        translator = GoogleTranslator(use_free_api=True)

        with pytest.raises(TranslationError):
            translator._translate_free_api("Hola")

    @patch("requests.post")
    def test_translate_paid_api_success(self, mock_post):
        """Test successful translation with paid API"""
        # Mock successful response
        mock_response = Mock()
        mock_response.json.return_value = {
            "data": {"translations": [{"translatedText": "Hello"}]}
        }
        mock_response.raise_for_status.return_value = None
        mock_post.return_value = mock_response

        translator = GoogleTranslator(api_key="test_key", use_free_api=False)
        result = translator._translate_paid_api("Hola")

        assert result == "Hello"
        mock_post.assert_called_once()

    def test_translate_paid_api_no_key(self):
        """Test paid API without API key"""
        translator = GoogleTranslator(use_free_api=False)

        with pytest.raises(TranslationError):
            translator._translate_paid_api("Hola")

    def test_permanent_error_detection(self):
        """Test permanent error detection"""
        translator = GoogleTranslator()

        # Test permanent errors
        auth_error = Exception("Invalid API key")
        assert translator._is_permanent_error(auth_error) is True

        bad_request_error = Exception("Bad request")
        assert translator._is_permanent_error(bad_request_error) is True

        # Test temporary errors
        network_error = Exception("Network timeout")
        assert translator._is_permanent_error(network_error) is False

    @patch("requests.get")
    def test_detect_language_free_api(self, mock_get):
        """Test language detection with free API"""
        # Mock response with language detection
        mock_response = Mock()
        mock_response.json.return_value = [None, None, "es"]
        mock_response.raise_for_status.return_value = None
        mock_get.return_value = mock_response

        translator = GoogleTranslator(use_free_api=True)
        result = translator.detect_language("Hola")

        assert result == "es"

    def test_get_supported_languages(self):
        """Test getting supported languages"""
        languages = get_supported_languages()

        assert isinstance(languages, dict)
        assert "en" in languages
        assert "zh" in languages
        assert "th" in languages
        assert languages["en"] == "English" or "English" in languages["en"]

    def test_validate_api_key_free(self):
        """Test API key validation for free API"""
        translator = GoogleTranslator(use_free_api=True)

        # Free API doesn't require key validation
        with patch.object(translator, "translate_text", return_value="translated"):
            assert translator.validate_api_key() is True

        with patch.object(translator, "translate_text", return_value="test"):
            assert translator.validate_api_key() is False


class TestBaseAPITranslator:
    """Test the BaseAPITranslator class"""

    class MockAPITranslator(BaseAPITranslator):
        """Mock implementation for testing"""

        def _translate_api_call(self, text: str) -> str:
            if "error" in text.lower():
                raise Exception("Mock API error")
            return f"translated_{text}"

    def test_initialization(self):
        """Test base API translator initialization"""
        translator = self.MockAPITranslator(api_key="test_key", rate_limit_requests=50)

        assert translator.api_key == "test_key"
        assert translator.rate_limit_requests == 50
        assert translator.rate_limit_window == 60

    def test_config_validation(self):
        """Test configuration validation"""
        # Valid config
        translator = self.MockAPITranslator()
        # Should not raise exception

        # Invalid config
        with pytest.raises(ConfigError):
            self.MockAPITranslator(rate_limit_requests=0)

    def test_rate_limiting(self):
        """Test rate limiting functionality"""
        translator = self.MockAPITranslator(rate_limit_requests=2, rate_limit_window=1)

        # First two requests should go through
        translator._check_rate_limit()
        translator._check_rate_limit()

        # Third request should be rate limited
        import time

        start_time = time.time()
        translator._check_rate_limit()
        end_time = time.time()

        # Should have waited (though might be very short in tests)
        assert end_time >= start_time

    def test_request_with_retry_success(self):
        """Test successful request with retry logic"""
        translator = self.MockAPITranslator()

        def mock_request(text):
            return f"success_{text}"

        result = translator._make_request_with_retry(mock_request, "test")
        assert result == "success_test"

    def test_request_with_retry_failure(self):
        """Test request failure with retry logic"""
        translator = self.MockAPITranslator(retry_count=2)

        def mock_request(text):
            raise Exception("API error")

        with pytest.raises(TranslationError):
            translator._make_request_with_retry(mock_request, "test")

    def test_request_with_permanent_error(self):
        """Test request with permanent error (no retry)"""
        translator = self.MockAPITranslator(retry_count=3)

        def mock_request(text):
            raise Exception("invalid api key")

        with pytest.raises(TranslationError):
            translator._make_request_with_retry(mock_request, "test")

    def test_translate_text_with_cache(self):
        """Test translate_text with caching"""
        translator = self.MockAPITranslator()

        # First translation - should call API
        result1 = translator.translate_text("hello")
        assert result1 == "translated_hello"

        # Second translation of same text - should use cache
        with patch.object(translator, "_translate_api_call") as mock_api:
            result2 = translator.translate_text("hello")
            assert result2 == "translated_hello"
            mock_api.assert_not_called()

    def test_api_info(self):
        """Test getting API information"""
        translator = self.MockAPITranslator(
            api_key="test_key", api_url="https://api.example.com"
        )

        info = translator.get_api_info()

        assert info["api_url"] == "https://api.example.com"
        assert info["has_api_key"] is True
        assert "rate_limit_requests" in info
        assert "current_request_count" in info

    def test_clear_rate_limit_history(self):
        """Test clearing rate limit history"""
        translator = self.MockAPITranslator()

        # Make some requests to build history
        translator._check_rate_limit()
        translator._check_rate_limit()

        assert len(translator._request_times) > 0

        # Clear history
        translator.clear_rate_limit_history()
        assert len(translator._request_times) == 0


def test_translator_factory():
    """Test translator factory function"""
    from offitrans.translators import get_translator

    # Test getting Google translator
    translator = get_translator("google", source_lang="zh", target_lang="en")
    assert isinstance(translator, GoogleTranslator)
    assert translator.source_lang == "zh"
    assert translator.target_lang == "en"

    # Test invalid translator type
    with pytest.raises(ValueError):
        get_translator("invalid_type")
