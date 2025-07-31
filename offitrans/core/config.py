"""
Configuration management for Offitrans

This module provides centralized configuration management for the translation system.
"""

import os
import json
import logging
from typing import Dict, Any, Optional, Union
from pathlib import Path
from dataclasses import dataclass, asdict

logger = logging.getLogger(__name__)


@dataclass
class TranslatorConfig:
    """Configuration for translator settings."""
    max_workers: int = 5
    timeout: int = 120
    retry_count: int = 3
    retry_delay: int = 2
    batch_size: int = 20
    api_key: Optional[str] = None
    api_url: Optional[str] = None


@dataclass
class CacheConfig:
    """Configuration for cache settings."""
    enabled: bool = True
    cache_file: str = "translation_cache.json"
    auto_save_interval: int = 10
    max_entries: int = 10000


@dataclass
class ProcessorConfig:
    """Configuration for file processor settings."""
    font_size_adjustment: float = 0.8
    preserve_formatting: bool = True
    image_protection: bool = True
    smart_column_width: bool = True


class Config:
    """
    Main configuration class for Offitrans.
    
    This class manages all configuration settings and provides methods
    to load from files, environment variables, and update settings.
    """
    
    def __init__(self, config_file: Optional[str] = None):
        """
        Initialize configuration.
        
        Args:
            config_file: Path to configuration file (optional)
        """
        # Default configurations
        self.translator = TranslatorConfig()
        self.cache = CacheConfig()
        self.processor = ProcessorConfig()
        
        # Additional settings
        self.debug: bool = False
        self.log_level: str = "INFO"
        self.supported_languages: Dict[str, str] = {
            'zh': 'Chinese',
            'en': 'English', 
            'th': 'Thai',
            'ja': 'Japanese',
            'ko': 'Korean',
            'fr': 'French',
            'de': 'German',
            'es': 'Spanish',
            'auto': 'Auto-detect'
        }
        
        # Load configuration from file if specified
        if config_file:
            self.load_from_file(config_file)
        
        # Override with environment variables
        self.load_from_env()
    
    def load_from_file(self, config_file: str) -> None:
        """
        Load configuration from JSON file.
        
        Args:
            config_file: Path to the configuration file
        """
        config_path = Path(config_file)
        
        if not config_path.exists():
            logger.warning(f"Configuration file not found: {config_file}")
            return
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # Update translator config
            if 'translator' in config_data:
                self._update_dataclass(self.translator, config_data['translator'])
            
            # Update cache config
            if 'cache' in config_data:
                self._update_dataclass(self.cache, config_data['cache'])
                
            # Update processor config
            if 'processor' in config_data:
                self._update_dataclass(self.processor, config_data['processor'])
            
            # Update other settings
            if 'debug' in config_data:
                self.debug = config_data['debug']
            if 'log_level' in config_data:
                self.log_level = config_data['log_level']
            if 'supported_languages' in config_data:
                self.supported_languages.update(config_data['supported_languages'])
                
            logger.info(f"Configuration loaded from: {config_file}")
            
        except Exception as e:
            logger.error(f"Failed to load configuration file {config_file}: {e}")
    
    def load_from_env(self) -> None:
        """Load configuration from environment variables."""
        # Translator settings
        if os.getenv('OFFITRANS_MAX_WORKERS'):
            self.translator.max_workers = int(os.getenv('OFFITRANS_MAX_WORKERS'))
        if os.getenv('OFFITRANS_TIMEOUT'):
            self.translator.timeout = int(os.getenv('OFFITRANS_TIMEOUT'))
        if os.getenv('OFFITRANS_RETRY_COUNT'):
            self.translator.retry_count = int(os.getenv('OFFITRANS_RETRY_COUNT'))
        if os.getenv('OFFITRANS_RETRY_DELAY'):
            self.translator.retry_delay = int(os.getenv('OFFITRANS_RETRY_DELAY'))
        if os.getenv('OFFITRANS_BATCH_SIZE'):
            self.translator.batch_size = int(os.getenv('OFFITRANS_BATCH_SIZE'))
        if os.getenv('OFFITRANS_API_KEY'):
            self.translator.api_key = os.getenv('OFFITRANS_API_KEY')
        if os.getenv('OFFITRANS_API_URL'):
            self.translator.api_url = os.getenv('OFFITRANS_API_URL')
        
        # Cache settings
        if os.getenv('OFFITRANS_CACHE_ENABLED'):
            self.cache.enabled = os.getenv('OFFITRANS_CACHE_ENABLED').lower() == 'true'
        if os.getenv('OFFITRANS_CACHE_FILE'):
            self.cache.cache_file = os.getenv('OFFITRANS_CACHE_FILE')
        if os.getenv('OFFITRANS_CACHE_AUTO_SAVE_INTERVAL'):
            self.cache.auto_save_interval = int(os.getenv('OFFITRANS_CACHE_AUTO_SAVE_INTERVAL'))
        if os.getenv('OFFITRANS_CACHE_MAX_ENTRIES'):
            self.cache.max_entries = int(os.getenv('OFFITRANS_CACHE_MAX_ENTRIES'))
        
        # Processor settings
        if os.getenv('OFFITRANS_FONT_SIZE_ADJUSTMENT'):
            self.processor.font_size_adjustment = float(os.getenv('OFFITRANS_FONT_SIZE_ADJUSTMENT'))
        if os.getenv('OFFITRANS_PRESERVE_FORMATTING'):
            self.processor.preserve_formatting = os.getenv('OFFITRANS_PRESERVE_FORMATTING').lower() == 'true'
        if os.getenv('OFFITRANS_IMAGE_PROTECTION'):
            self.processor.image_protection = os.getenv('OFFITRANS_IMAGE_PROTECTION').lower() == 'true'
        if os.getenv('OFFITRANS_SMART_COLUMN_WIDTH'):
            self.processor.smart_column_width = os.getenv('OFFITRANS_SMART_COLUMN_WIDTH').lower() == 'true'
        
        # General settings
        if os.getenv('OFFITRANS_DEBUG'):
            self.debug = os.getenv('OFFITRANS_DEBUG').lower() == 'true'
        if os.getenv('OFFITRANS_LOG_LEVEL'):
            self.log_level = os.getenv('OFFITRANS_LOG_LEVEL')
    
    def save_to_file(self, config_file: str) -> None:
        """
        Save current configuration to JSON file.
        
        Args:
            config_file: Path to save the configuration file
        """
        config_data = {
            'translator': asdict(self.translator),
            'cache': asdict(self.cache),
            'processor': asdict(self.processor),
            'debug': self.debug,
            'log_level': self.log_level,
            'supported_languages': self.supported_languages
        }
        
        try:
            config_path = Path(config_file)
            config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
            
            logger.info(f"Configuration saved to: {config_file}")
            
        except Exception as e:
            logger.error(f"Failed to save configuration to {config_file}: {e}")
    
    def update(self, **kwargs) -> None:
        """
        Update configuration with keyword arguments.
        
        Args:
            **kwargs: Configuration settings to update
        """
        for key, value in kwargs.items():
            if hasattr(self, key):
                setattr(self, key, value)
            elif hasattr(self.translator, key):
                setattr(self.translator, key, value)
            elif hasattr(self.cache, key):
                setattr(self.cache, key, value)
            elif hasattr(self.processor, key):
                setattr(self.processor, key, value)
            else:
                logger.warning(f"Unknown configuration key: {key}")
    
    def get_translator_kwargs(self) -> Dict[str, Any]:
        """
        Get translator configuration as keyword arguments.
        
        Returns:
            Dictionary of translator configuration
        """
        config = asdict(self.translator)
        config['supported_languages'] = self.supported_languages
        return config
    
    def get_cache_kwargs(self) -> Dict[str, Any]:
        """
        Get cache configuration as keyword arguments.
        
        Returns:
            Dictionary of cache configuration
        """
        return asdict(self.cache)
    
    def get_processor_kwargs(self) -> Dict[str, Any]:
        """
        Get processor configuration as keyword arguments.
        
        Returns:
            Dictionary of processor configuration
        """
        return asdict(self.processor)
    
    def _update_dataclass(self, target_obj, source_dict: Dict[str, Any]) -> None:
        """
        Update dataclass fields from dictionary.
        
        Args:
            target_obj: Target dataclass object
            source_dict: Source dictionary with new values
        """
        for key, value in source_dict.items():
            if hasattr(target_obj, key):
                setattr(target_obj, key, value)
            else:
                logger.warning(f"Unknown configuration field: {key}")
    
    def validate(self) -> bool:
        """
        Validate current configuration.
        
        Returns:
            True if configuration is valid, False otherwise
        """
        try:
            # Validate translator config
            if self.translator.max_workers <= 0:
                logger.error("max_workers must be positive")
                return False
            if self.translator.timeout <= 0:
                logger.error("timeout must be positive")
                return False
            if self.translator.retry_count < 0:
                logger.error("retry_count must be non-negative")
                return False
            if self.translator.batch_size <= 0:
                logger.error("batch_size must be positive")
                return False
            
            # Validate cache config
            if self.cache.auto_save_interval <= 0:
                logger.error("auto_save_interval must be positive")
                return False
            if self.cache.max_entries <= 0:
                logger.error("max_entries must be positive")
                return False
            
            # Validate processor config
            if self.processor.font_size_adjustment <= 0:
                logger.error("font_size_adjustment must be positive")
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"Configuration validation failed: {e}")
            return False
    
    def __str__(self) -> str:
        """String representation of configuration."""
        return f"Config(debug={self.debug}, log_level={self.log_level})"
    
    def __repr__(self) -> str:
        """Detailed string representation of configuration."""
        return (f"Config(translator={self.translator}, cache={self.cache}, "
                f"processor={self.processor}, debug={self.debug})")


# Global configuration instance
_global_config = Config()


def get_global_config() -> Config:
    """
    Get the global configuration instance.
    
    Returns:
        Global Config instance
    """
    return _global_config


def set_global_config(config: Config) -> None:
    """
    Set the global configuration instance.
    
    Args:
        config: New Config instance to use globally
    """
    global _global_config
    _global_config = config


def load_config_from_file(config_file: str) -> Config:
    """
    Load configuration from file and return new Config instance.
    
    Args:
        config_file: Path to configuration file
        
    Returns:
        New Config instance with loaded settings
    """
    return Config(config_file)