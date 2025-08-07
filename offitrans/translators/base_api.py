"""
Base API translator class for Offitrans

This module provides a base class for API-based translators with common
functionality like rate limiting, error handling, and retry logic.
"""

import time
import logging
from typing import Dict, Any, Optional
from abc import abstractmethod

from ..core.base import BaseTranslator
from ..core.cache import cached_translation
from ..exceptions.errors import TranslationError, ConfigError

logger = logging.getLogger(__name__)


class BaseAPITranslator(BaseTranslator):
    """
    Base class for API-based translators.

    This class provides common functionality for translators that use
    external APIs, including rate limiting, error handling, and caching.
    """

    def __init__(
        self,
        api_key: Optional[str] = None,
        api_url: Optional[str] = None,
        rate_limit_requests: int = 100,
        rate_limit_window: int = 60,
        proxies: Optional[Dict[str, str]] = None,
        **kwargs,
    ):
        """
        Initialize API-based translator.

        Args:
            api_key: API key for the translation service
            api_url: Base URL for the API (optional)
            rate_limit_requests: Maximum requests per window (default: 100)
            rate_limit_window: Rate limit window in seconds (default: 60)
            proxies: Proxy configuration dict (e.g., {'http': 'http://127.0.0.1:7890', 'https': 'http://127.0.0.1:7890'})
            **kwargs: Additional arguments passed to BaseTranslator
        """
        super().__init__(**kwargs)

        self.api_key = api_key
        self.api_url = api_url
        self.rate_limit_requests = rate_limit_requests
        self.rate_limit_window = rate_limit_window
        self.proxies = proxies

        # Rate limiting tracking
        self._request_times = []

        # Validate configuration
        self._validate_config()

    def _validate_config(self) -> None:
        """
        Validate translator configuration.

        Raises:
            ConfigError: If configuration is invalid
        """
        if not self.api_key:
            logger.warning("No API key provided - some functionality may be limited")

        if self.rate_limit_requests <= 0:
            raise ConfigError("rate_limit_requests must be positive")

        if self.rate_limit_window <= 0:
            raise ConfigError("rate_limit_window must be positive")

    def _check_rate_limit(self) -> None:
        """
        Check and enforce rate limiting.

        Raises:
            TranslationError: If rate limit is exceeded
        """
        current_time = time.time()

        # Remove old requests outside the window
        self._request_times = [
            t for t in self._request_times if current_time - t < self.rate_limit_window
        ]

        # Check if we're at the limit
        if len(self._request_times) >= self.rate_limit_requests:
            oldest_request = min(self._request_times)
            wait_time = self.rate_limit_window - (current_time - oldest_request)

            if wait_time > 0:
                logger.warning(
                    f"Rate limit reached. Waiting {wait_time:.1f} seconds..."
                )
                time.sleep(wait_time)

        # Record this request
        self._request_times.append(current_time)

    def _make_request_with_retry(self, request_func, *args, **kwargs) -> Any:
        """
        Make an API request with retry logic.

        Args:
            request_func: Function that makes the API request
            *args: Arguments for the request function
            **kwargs: Keyword arguments for the request function

        Returns:
            Response from the API

        Raises:
            TranslationError: If all retry attempts fail
        """
        last_exception = None

        for attempt in range(self.retry_count + 1):
            try:
                # Check rate limit before making request
                self._check_rate_limit()

                # Make the request
                response = request_func(*args, **kwargs)

                # Update statistics
                self._update_stats(success=True)
                return response

            except Exception as e:
                last_exception = e
                logger.error(f"API request attempt {attempt + 1} failed: {e}")

                # Update statistics
                self._update_stats(success=False)

                # Don't retry on certain errors
                if self._is_permanent_error(e):
                    logger.error("Permanent error detected, not retrying")
                    break

                # Wait before retry (exponential backoff)
                if attempt < self.retry_count:
                    wait_time = self.retry_delay * (2**attempt)
                    logger.info(f"Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)

        # All attempts failed
        error_msg = f"API request failed after {self.retry_count + 1} attempts"
        if last_exception:
            error_msg += f": {last_exception}"

        raise TranslationError(error_msg) from last_exception

    def _is_permanent_error(self, error: Exception) -> bool:
        """
        Check if an error is permanent (shouldn't retry).

        Args:
            error: Exception to check

        Returns:
            True if error is permanent, False otherwise
        """
        # Override in subclasses to define permanent errors
        # Common permanent errors: authentication failures, invalid requests
        error_str = str(error).lower()
        permanent_keywords = [
            "authentication",
            "unauthorized",
            "forbidden",
            "invalid api key",
            "bad request",
            "not found",
            "method not allowed",
        ]

        return any(keyword in error_str for keyword in permanent_keywords)

    @abstractmethod
    def _translate_api_call(self, text: str) -> str:
        """
        Make the actual API call to translate text.

        This method should be implemented by subclasses to handle
        the specific API call format.

        Args:
            text: Text to translate

        Returns:
            Translated text

        Raises:
            TranslationError: If API call fails
        """
        pass

    @cached_translation()
    def translate_text(self, text: str) -> str:
        """
        Translate a single text string using the API.

        Args:
            text: Text to translate

        Returns:
            Translated text

        Raises:
            TranslationError: If translation fails
        """
        if not text or not text.strip():
            return text

        try:
            # Use the request wrapper for retry logic
            result = self._make_request_with_retry(self._translate_api_call, text)
            return result

        except Exception as e:
            logger.error(f"Translation failed for text: {text[:50]}... Error: {e}")
            raise TranslationError(f"Translation failed: {e}") from e

    def validate_api_key(self) -> bool:
        """
        Validate the API key by making a test request.

        Returns:
            True if API key is valid, False otherwise
        """
        try:
            # Test with a simple translation
            test_text = "test"
            result = self.translate_text(test_text)
            return result != test_text  # Should be translated

        except Exception as e:
            logger.error(f"API key validation failed: {e}")
            return False

    def get_api_info(self) -> Dict[str, Any]:
        """
        Get information about the API configuration.

        Returns:
            Dictionary with API information
        """
        return {
            "api_url": self.api_url,
            "has_api_key": bool(self.api_key),
            "rate_limit_requests": self.rate_limit_requests,
            "rate_limit_window": self.rate_limit_window,
            "current_request_count": len(self._request_times),
        }

    def clear_rate_limit_history(self) -> None:
        """Clear rate limiting history."""
        self._request_times.clear()
        logger.info("Rate limit history cleared")
