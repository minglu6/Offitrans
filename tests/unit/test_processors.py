"""
Unit tests for processor classes
"""

import pytest
from unittest.mock import Mock, patch, MagicMock
from pathlib import Path

from offitrans.processors.base import BaseProcessor
from offitrans.processors import get_processor, get_processor_by_extension
from offitrans.exceptions.errors import ProcessorError


class TestBaseProcessor:
    """Test the BaseProcessor class"""

    class MockProcessor(BaseProcessor):
        """Mock implementation for testing"""

        def extract_text(self, file_path: str):
            return [
                {"text": "Hello world", "type": "test"},
                {"text": "Hello world", "type": "test"},
            ]

        def translate_and_save(
            self, file_path: str, output_path: str, target_language: str = "en"
        ):
            return True

    def test_initialization(self):
        """Test base processor initialization"""
        processor = self.MockProcessor()

        assert processor.translator is not None
        assert processor.preserve_formatting is True
        assert processor.image_protection is True
        assert hasattr(processor, "stats")

    def test_initialization_with_custom_translator(self):
        """Test initialization with custom translator"""
        mock_translator = Mock()
        processor = self.MockProcessor(translator=mock_translator)

        assert processor.translator is mock_translator

    def test_file_validation_exists(self, temp_dir):
        """Test file validation for existing file"""
        processor = self.MockProcessor()

        # Create a test file
        test_file = temp_dir / "test.txt"
        test_file.write_text("test content")

        # Mock supports_file_type to return True
        with patch.object(processor, "supports_file_type", return_value=True):
            assert processor.validate_file(str(test_file)) is True

    def test_file_validation_not_exists(self):
        """Test file validation for non-existent file"""
        processor = self.MockProcessor()

        assert processor.validate_file("nonexistent.txt") is False

    def test_file_validation_too_large(self, temp_dir):
        """Test file validation for oversized file"""
        processor = self.MockProcessor()

        # Create a test file
        test_file = temp_dir / "large.txt"
        test_file.write_text("x" * 1000)

        # Mock file size check
        with patch.object(Path, "stat") as mock_stat:
            mock_stat.return_value.st_size = 200 * 1024 * 1024  # 200MB
            assert processor.validate_file(str(test_file)) is False

    def test_preprocess_texts(self):
        """Test text preprocessing"""
        processor = self.MockProcessor()

        texts = [
            "Hello world",
            "Hello world",
            "123",  # Should be filtered out
            "test@email.com",  # Should be filtered out
            "Hello world",  # Duplicate
        ]

        unique_texts, metadata = processor.preprocess_texts(texts)

        assert len(unique_texts) <= len(texts)  # Should be deduplicated
        assert metadata["original_count"] == 5
        assert metadata["unique_count"] <= 5
        assert "text_to_indices" in metadata

    def test_translate_texts(self):
        """Test text translation"""
        processor = self.MockProcessor()

        # Mock translator
        mock_translator = Mock()
        mock_translator.translate_text_batch.return_value = ["Hello", "World"]
        processor.translator = mock_translator

        texts = ["Hello", "World"]
        result = processor.translate_texts(texts, "en")

        assert result == ["Hello", "World"]
        mock_translator.translate_text_batch.assert_called_once_with(texts)

    def test_postprocess_translations(self):
        """Test translation post-processing"""
        processor = self.MockProcessor()

        original_texts = ["Hello", "World", "123"]
        translated_texts = ["Hello", "World"]
        metadata = {
            "text_to_indices": {"Hello": [0], "World": [1]},
            "non_translatable_texts": ["123"],
        }

        result = processor.postprocess_translations(
            original_texts, translated_texts, metadata
        )

        assert len(result) == 3
        assert result[2] == "123"  # Non-translatable should remain unchanged

    def test_process_file_success(self, temp_dir):
        """Test successful file processing"""
        processor = self.MockProcessor()

        # Create test files
        input_file = temp_dir / "input.txt"
        output_file = temp_dir / "output.txt"
        input_file.write_text("test content")

        # Mock validation to return True
        with patch.object(processor, "validate_file", return_value=True):
            result = processor.process_file(str(input_file), str(output_file))

        assert result is True
        assert processor.stats["total_files_processed"] == 1
        assert processor.stats["successful_files"] == 1

    def test_process_file_validation_failure(self):
        """Test file processing with validation failure"""
        processor = self.MockProcessor()

        result = processor.process_file("nonexistent.txt", "output.txt")

        assert result is False
        assert processor.stats["failed_files"] == 1

    def test_get_stats(self):
        """Test statistics retrieval"""
        processor = self.MockProcessor()

        stats = processor.get_stats()

        assert isinstance(stats, dict)
        assert "total_files_processed" in stats
        assert "successful_files" in stats
        assert "failed_files" in stats

    def test_reset_stats(self):
        """Test statistics reset"""
        processor = self.MockProcessor()

        # Simulate some processing
        processor.stats["total_files_processed"] = 5
        processor.stats["successful_files"] = 3

        processor.reset_stats()

        assert processor.stats["total_files_processed"] == 0
        assert processor.stats["successful_files"] == 0


@pytest.mark.requires_openpyxl
class TestExcelProcessor:
    """Test Excel processor (requires openpyxl)"""

    def test_supports_file_type(self):
        """Test file type support check"""
        try:
            from offitrans.processors.excel import ExcelProcessor

            processor = ExcelProcessor()

            assert processor.supports_file_type("test.xlsx") is True
            assert processor.supports_file_type("test.xls") is True
            assert processor.supports_file_type("test.xlsm") is True
            assert processor.supports_file_type("test.txt") is False
        except ImportError:
            pytest.skip("openpyxl not available")

    def test_initialization_without_openpyxl(self):
        """Test Excel processor initialization without openpyxl"""
        with patch.dict("sys.modules", {"openpyxl": None}):
            from offitrans.processors.excel import OPENPYXL_AVAILABLE

            if not OPENPYXL_AVAILABLE:
                from offitrans.exceptions.errors import ExcelProcessorError

                with pytest.raises(ExcelProcessorError):
                    from offitrans.processors.excel import ExcelProcessor

                    ExcelProcessor()


@pytest.mark.requires_docx
class TestWordProcessor:
    """Test Word processor (requires python-docx)"""

    def test_supports_file_type(self):
        """Test file type support check"""
        try:
            from offitrans.processors.word import WordProcessor

            processor = WordProcessor()

            assert processor.supports_file_type("test.docx") is True
            assert processor.supports_file_type("test.doc") is True
            assert processor.supports_file_type("test.txt") is False
        except ImportError:
            pytest.skip("python-docx not available")


@pytest.mark.requires_pptx
class TestPowerPointProcessor:
    """Test PowerPoint processor (requires python-pptx)"""

    def test_supports_file_type(self):
        """Test file type support check"""
        try:
            from offitrans.processors.powerpoint import PowerPointProcessor

            processor = PowerPointProcessor()

            assert processor.supports_file_type("test.pptx") is True
            assert processor.supports_file_type("test.ppt") is True
            assert processor.supports_file_type("test.txt") is False
        except ImportError:
            pytest.skip("python-pptx not available")


class TestPDFProcessor:
    """Test PDF processor"""

    def test_supports_file_type(self):
        """Test file type support check"""
        try:
            from offitrans.processors.pdf import PDFProcessor

            processor = PDFProcessor()

            assert processor.supports_file_type("test.pdf") is True
            assert processor.supports_file_type("test.txt") is False
        except ImportError:
            pytest.skip("PyPDF2 not available")


class TestProcessorFactory:
    """Test processor factory functions"""

    def test_get_processor_by_type(self):
        """Test getting processor by type"""
        # Test with a type that should always be available
        processor = get_processor("pdf")
        assert processor is not None

    def test_get_processor_invalid_type(self):
        """Test getting processor with invalid type"""
        with pytest.raises(ValueError):
            get_processor("invalid_type")

    def test_get_processor_by_extension(self, temp_dir):
        """Test getting processor by file extension"""
        test_file = temp_dir / "test.pdf"
        test_file.touch()

        processor = get_processor_by_extension(str(test_file))
        assert processor is not None

    def test_get_processor_by_extension_invalid(self, temp_dir):
        """Test getting processor with invalid extension"""
        test_file = temp_dir / "test.invalid"
        test_file.touch()

        with pytest.raises(ValueError):
            get_processor_by_extension(str(test_file))
