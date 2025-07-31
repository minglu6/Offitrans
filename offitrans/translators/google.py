"""
Google Translate API implementation for Offitrans

This module provides integration with Google Cloud Translation API.
"""

import html
import requests
import logging
from typing import Dict, Any, Optional

from .base_api import BaseAPITranslator
from ..exceptions.errors import TranslationError

logger = logging.getLogger(__name__)


def get_supported_languages() -> Dict[str, str]:
    """
    Get list of supported languages for Google Translate.
    
    Returns:
        Dictionary mapping language codes to language names
    """
    return {
        'zh': 'Chinese (中文)',
        'en': 'English',
        'th': 'ไทย (Thai)',
        'ja': '日本語 (Japanese)',
        'ko': '한국어 (Korean)',
        'fr': 'Français (French)',
        'de': 'Deutsch (German)',
        'es': 'Español (Spanish)',
        'ar': 'العربية (Arabic)',
        'ru': 'Русский (Russian)',
        'pt': 'Português (Portuguese)',
        'it': 'Italiano (Italian)',
        'hi': 'हिन्दी (Hindi)',
        'auto': 'Auto-detect'
    }


class GoogleTranslator(BaseAPITranslator):
    """
    Google Cloud Translation API implementation.
    
    This translator uses the Google Cloud Translation API to translate text.
    It supports both the free API and the paid Cloud Translation API.
    """
    
    def __init__(self, 
                 api_key: Optional[str] = None,
                 use_free_api: bool = True,
                 **kwargs):
        """
        Initialize Google Translator.
        
        Args:
            api_key: Google Cloud API key (optional for free API)
            use_free_api: Whether to use the free Google Translate API (default: True)
            **kwargs: Additional arguments passed to BaseAPITranslator
        """
        # Set default API URL based on API type
        if 'api_url' not in kwargs:
            if use_free_api:
                kwargs['api_url'] = "https://translate.googleapis.com/translate_a/single"
            else:
                kwargs['api_url'] = "https://translation.googleapis.com/language/translate/v2"
        
        super().__init__(api_key=api_key, **kwargs)
        
        self.use_free_api = use_free_api
        
        # Update supported languages
        self.supported_languages.update(get_supported_languages())
        
        # Set reasonable rate limits for Google API
        if not hasattr(self, 'rate_limit_requests'):
            self.rate_limit_requests = 100 if use_free_api else 1000
        if not hasattr(self, 'rate_limit_window'):
            self.rate_limit_window = 60
    
    def _translate_api_call(self, text: str) -> str:
        """
        Make the actual Google Translate API call.
        
        Args:
            text: Text to translate
            
        Returns:
            Translated text
            
        Raises:
            TranslationError: If API call fails
        """
        if self.use_free_api:
            return self._translate_free_api(text)
        else:
            return self._translate_paid_api(text)
    
    def _translate_free_api(self, text: str) -> str:
        """
        Use the free Google Translate API.
        
        Args:
            text: Text to translate
            
        Returns:
            Translated text
        """
        try:
            # Free API parameters
            params = {
                'client': 'gtx',
                'sl': self.source_lang if self.source_lang != 'auto' else 'auto',
                'tl': self.target_lang,
                'dt': 't',
                'q': text
            }
            
            response = requests.get(
                self.api_url,
                params=params,
                timeout=self.timeout,
                headers={'User-Agent': 'Mozilla/5.0 (compatible; Offitrans/1.0)'}
            )
            response.raise_for_status()
            
            # Parse the response
            result = response.json()
            if result and len(result) > 0 and len(result[0]) > 0:
                translated_text = result[0][0][0]
                # Decode HTML entities
                translated_text = html.unescape(translated_text)
                return translated_text
            else:
                raise TranslationError(f"Empty response from Google Translate API")
                
        except requests.exceptions.RequestException as e:
            raise TranslationError(f"Request failed: {e}") from e
        except (IndexError, KeyError, TypeError) as e:
            raise TranslationError(f"Failed to parse API response: {e}") from e
    
    def _translate_paid_api(self, text: str) -> str:
        """
        Use the paid Google Cloud Translation API.
        
        Args:
            text: Text to translate
            
        Returns:
            Translated text
        """
        if not self.api_key:
            raise TranslationError("API key required for Google Cloud Translation API")
        
        try:
            # Language mapping for the paid API
            lang_map = {
                'en': 'en',
                'zh': 'zh',
                'th': 'th',
                'ja': 'ja',
                'ko': 'ko',
                'fr': 'fr',
                'de': 'de',
                'es': 'es',
                'ar': 'ar',
                'ru': 'ru',
                'pt': 'pt',
                'it': 'it',
                'hi': 'hi'
            }
            
            target_lang_code = lang_map.get(self.target_lang, self.target_lang)
            source_lang_code = lang_map.get(self.source_lang, self.source_lang)
            
            # Paid API parameters
            params = {
                'key': self.api_key,
                'q': text,
                'target': target_lang_code,
                'format': 'text'
            }
            
            # Add source language if not auto-detect
            if source_lang_code != 'auto':
                params['source'] = source_lang_code
            
            response = requests.post(
                self.api_url,
                data=params,
                timeout=self.timeout,
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            response.raise_for_status()
            
            # Parse the response
            result_json = response.json()
            
            if ('data' in result_json and
                'translations' in result_json['data'] and
                len(result_json['data']['translations']) > 0):
                
                translated_text = result_json['data']['translations'][0]['translatedText']
                # Decode HTML entities
                translated_text = html.unescape(translated_text)
                return translated_text
            else:
                raise TranslationError("Invalid response format from Google Cloud API")
                
        except requests.exceptions.RequestException as e:
            raise TranslationError(f"Request failed: {e}") from e
        except (KeyError, IndexError, TypeError) as e:
            raise TranslationError(f"Failed to parse API response: {e}") from e
    
    def _is_permanent_error(self, error: Exception) -> bool:
        """
        Check if an error is permanent for Google Translate API.
        
        Args:
            error: Exception to check
            
        Returns:
            True if error is permanent, False otherwise
        """
        # Call parent method first
        if super()._is_permanent_error(error):
            return True
        
        # Google-specific permanent errors
        error_str = str(error).lower()
        google_permanent_errors = [
            'invalid api key',
            'api key not valid',
            'daily limit exceeded',
            'user rate limit exceeded',
            'bad request',
            'invalid request'
        ]
        
        return any(error_phrase in error_str for error_phrase in google_permanent_errors)
    
    def detect_language(self, text: str) -> str:
        """
        Detect the language of the given text using Google API.
        
        Args:
            text: Text to analyze
            
        Returns:
            Detected language code
        """
        if not text or not text.strip():
            return 'unknown'
        
        try:
            if self.use_free_api:
                # Free API doesn't have dedicated detection, use translate with auto
                params = {
                    'client': 'gtx',
                    'sl': 'auto',
                    'tl': 'en',  # Translate to English to get source detection
                    'dt': 't',
                    'q': text[:100]  # Use first 100 chars for detection
                }
                
                response = requests.get(
                    self.api_url,
                    params=params,
                    timeout=self.timeout
                )
                response.raise_for_status()
                
                result = response.json()
                # The detected language is in result[2] for free API
                if len(result) > 2 and result[2]:
                    return result[2]
                    
            else:
                # Use paid API's detection endpoint
                if not self.api_key:
                    return 'unknown'
                
                detect_url = "https://translation.googleapis.com/language/translate/v2/detect"
                params = {
                    'key': self.api_key,
                    'q': text[:1000]  # Limit text length for detection
                }
                
                response = requests.post(detect_url, data=params, timeout=self.timeout)
                response.raise_for_status()
                
                result = response.json()
                if ('data' in result and 'detections' in result['data'] and
                    len(result['data']['detections']) > 0 and
                    len(result['data']['detections'][0]) > 0):
                    return result['data']['detections'][0][0]['language']
            
            return 'unknown'
            
        except Exception as e:
            logger.error(f"Language detection failed: {e}")
            return 'unknown'
    
    def get_supported_languages_from_api(self) -> Dict[str, str]:
        """
        Get supported languages directly from Google API.
        
        Returns:
            Dictionary of supported language codes and names
        """
        if not self.use_free_api and self.api_key:
            try:
                url = "https://translation.googleapis.com/language/translate/v2/languages"
                params = {
                    'key': self.api_key,
                    'target': 'en'  # Get language names in English
                }
                
                response = requests.get(url, params=params, timeout=self.timeout)
                response.raise_for_status()
                
                result = response.json()
                if 'data' in result and 'languages' in result['data']:
                    languages = {}
                    for lang in result['data']['languages']:
                        languages[lang['language']] = lang['name']
                    return languages
                    
            except Exception as e:
                logger.error(f"Failed to get languages from API: {e}")
        
        # Fall back to static list
        return get_supported_languages()
    
    def validate_api_key(self) -> bool:
        """
        Validate the Google API key.
        
        Returns:
            True if API key is valid, False otherwise
        """
        if self.use_free_api:
            # Free API doesn't require key validation
            try:
                result = self.translate_text("test")
                return bool(result and result != "test")
            except Exception:
                return False
        else:
            # Use parent validation for paid API
            return super().validate_api_key()
    
    def __str__(self) -> str:
        """String representation of the translator."""
        api_type = "Free" if self.use_free_api else "Paid"
        return f"GoogleTranslator({api_type}, {self.source_lang} -> {self.target_lang})"