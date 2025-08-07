"""
Base processor class for Offitrans

This module provides the abstract base class for all file processors.
"""

from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import logging

from ..core.config import Config, get_global_config
from ..translators import GoogleTranslator
from ..exceptions.errors import ProcessorError, FileError

logger = logging.getLogger(__name__)


class BaseProcessor(ABC):
    """
    Abstract base class for all file processors.

    This class defines the common interface and functionality that all
    file processor implementations must provide.
    """

    def __init__(self, translator=None, config: Optional[Config] = None, **kwargs):
        """
        Initialize the base processor.

        Args:
            translator: Translator instance to use (default: GoogleTranslator)
            config: Configuration instance (default: global config)
            **kwargs: Additional keyword arguments
        """
        # Use provided config or global config
        self.config = config or get_global_config()

        # Initialize translator
        if translator is None:
            translator_config = self.config.get_translator_kwargs()
            self.translator = GoogleTranslator(**translator_config)
        else:
            self.translator = translator

        # Processor settings from config
        self.preserve_formatting = self.config.processor.preserve_formatting
        self.image_protection = self.config.processor.image_protection
        self.font_size_adjustment = self.config.processor.font_size_adjustment

        # Statistics
        self.stats = {
            "total_files_processed": 0,
            "successful_files": 0,
            "failed_files": 0,
            "total_texts_translated": 0,
            "total_chars_translated": 0,
        }

        # Apply additional kwargs
        for key, value in kwargs.items():
            if hasattr(self, key):
                setattr(self, key, value)

    @abstractmethod
    def extract_text(self, file_path: str) -> List[Dict[str, Any]]:
        """
        Extract text content from the file.

        Args:
            file_path: Path to the input file

        Returns:
            List of dictionaries containing text and metadata

        Raises:
            ProcessorError: If text extraction fails
        """
        pass

    @abstractmethod
    def translate_and_save(
        self, file_path: str, output_path: str, target_language: str = "en"
    ) -> bool:
        """
        Translate the file and save to output path.

        Args:
            file_path: Path to the input file
            output_path: Path for the output file
            target_language: Target language code

        Returns:
            True if successful, False otherwise

        Raises:
            ProcessorError: If translation or saving fails
        """
        pass

    def validate_file(self, file_path: str) -> bool:
        """
        Validate if the file can be processed.

        Args:
            file_path: Path to the file to validate

        Returns:
            True if file is valid, False otherwise
        """
        try:
            file_path = Path(file_path)

            # Check if file exists
            if not file_path.exists():
                logger.error(f"File does not exist: {file_path}")
                return False

            # Check if it's a file (not a directory)
            if not file_path.is_file():
                logger.error(f"Path is not a file: {file_path}")
                return False

            # Check file size (avoid processing very large files)
            file_size = file_path.stat().st_size
            max_size = 100 * 1024 * 1024  # 100MB limit
            if file_size > max_size:
                logger.error(f"File too large: {file_size} bytes (max: {max_size})")
                return False

            # Check if file extension is supported
            if not self.supports_file_type(file_path):
                logger.error(f"Unsupported file type: {file_path.suffix}")
                return False

            return True

        except Exception as e:
            logger.error(f"File validation failed: {e}")
            return False

    def supports_file_type(self, file_path: str) -> bool:
        """
        Check if the processor supports the file type.

        Args:
            file_path: Path to the file

        Returns:
            True if file type is supported, False otherwise
        """
        # Override in subclasses to define supported extensions
        return True

    def preprocess_texts(self, texts: List[str]) -> Tuple[List[str], Dict[str, Any]]:
        """
        Preprocess texts before translation (filtering, deduplication, etc.).

        Args:
            texts: List of original texts

        Returns:
            Tuple of (processed_texts, metadata)
        """
        from ..core.utils import filter_translatable_texts, deduplicate_texts

        # Filter translatable texts
        translatable, non_translatable = filter_translatable_texts(texts)

        # Deduplicate texts
        unique_texts, text_to_indices = deduplicate_texts(translatable)

        metadata = {
            "original_count": len(texts),
            "translatable_count": len(translatable),
            "non_translatable_count": len(non_translatable),
            "unique_count": len(unique_texts),
            "non_translatable_texts": non_translatable,
            "text_to_indices": text_to_indices,
        }

        logger.info(
            f"Text preprocessing: {len(texts)} → {len(unique_texts)} unique translatable texts"
        )

        return unique_texts, metadata

    def translate_texts(
        self, texts: List[str], target_language: str = "en"
    ) -> List[str]:
        """
        Translate a list of texts.

        Args:
            texts: List of texts to translate
            target_language: Target language code

        Returns:
            List of translated texts
        """
        if not texts:
            return []

        # Update translator target language
        self.translator.target_lang = target_language

        # Use batch translation for efficiency
        translated = self.translator.translate_text_batch(texts)

        # Update statistics
        self.stats["total_texts_translated"] += len(texts)
        self.stats["total_chars_translated"] += sum(len(text) for text in texts)

        return translated

    def postprocess_translations(
        self,
        original_texts: List[str],
        translated_texts: List[str],
        metadata: Dict[str, Any],
    ) -> List[str]:
        """
        Post-process translations (mapping back to original structure).

        Args:
            original_texts: Original text list
            translated_texts: Translated text list
            metadata: Metadata from preprocessing

        Returns:
            List of translations mapped back to original structure
        """
        # Reconstruct full translation list
        text_to_indices = metadata.get("text_to_indices", {})
        non_translatable_texts = metadata.get("non_translatable_texts", [])

        # Create translation mapping
        translation_map = {}
        for i, (original, translated) in enumerate(
            zip(
                [text for text in original_texts if text in text_to_indices],
                translated_texts,
            )
        ):
            translation_map[original] = translated

        # Map back to original structure
        result = []
        for original_text in original_texts:
            if original_text in translation_map:
                result.append(translation_map[original_text])
            elif original_text in non_translatable_texts:
                result.append(original_text)  # Keep original for non-translatable
            else:
                result.append(original_text)  # Fallback to original

        return result

    def process_file(
        self, input_path: str, output_path: str, target_language: str = "en"
    ) -> bool:
        """
        High-level method to process a file (extract → translate → save).

        Args:
            input_path: Path to input file
            output_path: Path to output file
            target_language: Target language code

        Returns:
            True if successful, False otherwise
        """
        try:
            # Validate input file
            if not self.validate_file(input_path):
                return False

            # Update statistics
            self.stats["total_files_processed"] += 1

            # Delegate to implementation-specific method
            success = self.translate_and_save(input_path, output_path, target_language)

            # Update statistics
            if success:
                self.stats["successful_files"] += 1
                logger.info(f"Successfully processed: {input_path} → {output_path}")
            else:
                self.stats["failed_files"] += 1
                logger.error(f"Failed to process: {input_path}")

            return success

        except Exception as e:
            self.stats["failed_files"] += 1
            logger.error(f"Error processing file {input_path}: {e}")
            return False

    def get_stats(self) -> Dict[str, Any]:
        """
        Get processing statistics.

        Returns:
            Dictionary containing processing statistics
        """
        return self.stats.copy()

    def reset_stats(self) -> None:
        """Reset processing statistics."""
        self.stats = {
            "total_files_processed": 0,
            "successful_files": 0,
            "failed_files": 0,
            "total_texts_translated": 0,
            "total_chars_translated": 0,
        }

    def __str__(self) -> str:
        """String representation of the processor."""
        return f"{self.__class__.__name__}(translator={self.translator})"

    def __repr__(self) -> str:
        """Detailed string representation of the processor."""
        return (
            f"{self.__class__.__name__}("
            f"translator={repr(self.translator)}, "
            f"preserve_formatting={self.preserve_formatting}, "
            f"image_protection={self.image_protection})"
        )
