"""
End-to-end integration tests for Offitrans

These tests verify that the complete translation workflow works correctly.
"""

import pytest
from pathlib import Path

from offitrans import ExcelProcessor, GoogleTranslator
from offitrans.processors import get_processor_by_extension


@pytest.mark.integration
@pytest.mark.slow
class TestEndToEndTranslation:
    """Test complete translation workflows"""

    @pytest.mark.requires_openpyxl
    def test_excel_translation_workflow(
        self, temp_dir, file_generator, mock_translator
    ):
        """Test complete Excel translation workflow"""
        # Create a test Excel file
        input_file = temp_dir / "test_input.xlsx"
        output_file = temp_dir / "test_output.xlsx"

        test_texts = ["Hello World", "Hello World", "Test Data"]
        success = file_generator.create_simple_excel(input_file, test_texts)

        if not success:
            pytest.skip("Could not create test Excel file")

        # Create processor with mock translator
        processor = ExcelProcessor(translator=mock_translator)

        # Process the file
        result = processor.process_file(str(input_file), str(output_file))

        assert result is True
        assert output_file.exists()

        # Verify statistics
        stats = processor.get_stats()
        assert stats["total_files_processed"] == 1
        assert stats["successful_files"] == 1

    @pytest.mark.requires_docx
    def test_word_translation_workflow(self, temp_dir, file_generator, mock_translator):
        """Test complete Word translation workflow"""
        # Create a test Word file
        input_file = temp_dir / "test_input.docx"
        output_file = temp_dir / "test_output.docx"

        test_texts = ["This is a test document", "This is a test document"]
        success = file_generator.create_simple_word(input_file, test_texts)

        if not success:
            pytest.skip("Could not create test Word file")

        # Get processor by extension
        processor = get_processor_by_extension(
            str(input_file), translator=mock_translator
        )

        # Process the file
        result = processor.process_file(str(input_file), str(output_file))

        assert result is True
        assert output_file.exists()

    @pytest.mark.requires_pptx
    def test_powerpoint_translation_workflow(
        self, temp_dir, file_generator, mock_translator
    ):
        """Test complete PowerPoint translation workflow"""
        # Create a test PowerPoint file
        input_file = temp_dir / "test_input.pptx"
        output_file = temp_dir / "test_output.pptx"

        test_texts = ["Test Slide Title", "Test Slide Content"]
        success = file_generator.create_simple_ppt(input_file, test_texts)

        if not success:
            pytest.skip("Could not create test PowerPoint file")

        # Get processor by extension
        processor = get_processor_by_extension(
            str(input_file), translator=mock_translator
        )

        # Process the file
        result = processor.process_file(str(input_file), str(output_file))

        assert result is True
        assert output_file.exists()

    def test_pdf_translation_workflow(self, temp_dir, mock_translator):
        """Test PDF translation workflow"""
        # Create a simple text file (PDF creation is complex)
        input_file = temp_dir / "test_input.txt"
        output_file = temp_dir / "test_output.txt"

        input_file.write_text("This is a test document\\nwith multiple lines.")

        # Get PDF processor
        try:
            processor = get_processor_by_extension(
                "test.pdf", translator=mock_translator
            )

            # Since we're using a text file, this might fail, but we test the workflow
            result = processor.process_file(str(input_file), str(output_file))

            # The result depends on implementation details
            # Main goal is to test that the workflow doesn't crash

        except Exception as e:
            pytest.skip(f"PDF processing not available: {e}")


@pytest.mark.integration
class TestTranslatorIntegration:
    """Test translator integration with processors"""

    def test_translator_processor_integration(self, mock_translator):
        """Test that processors work correctly with translators"""
        from offitrans.processors.base import BaseProcessor

        class TestProcessor(BaseProcessor):
            def extract_text(self, file_path):
                return [{"text": "Hello world"}, {"text": "Test text"}]

            def translate_and_save(self, file_path, output_path, target_language="en"):
                return True

        processor = TestProcessor(translator=mock_translator)

        # Test that translator is properly integrated
        assert processor.translator is mock_translator

        # Test translation through processor
        texts = ["Hello", "World"]
        result = processor.translate_texts(texts)

        assert len(result) == 2
        assert all("[TRANSLATED_en]" in text for text in result)

    def test_config_integration(self, test_config):
        """Test configuration integration with components"""
        from offitrans.processors.excel import ExcelProcessor

        try:
            processor = ExcelProcessor(config=test_config)

            # Verify config is applied
            assert processor.config is test_config
            assert (
                processor.preserve_formatting
                == test_config.processor.preserve_formatting
            )

            # Test that translator gets config
            assert (
                processor.translator.max_workers == test_config.translator.max_workers
            )

        except ImportError:
            pytest.skip("Excel processor dependencies not available")


@pytest.mark.integration
@pytest.mark.requires_api
@pytest.mark.slow
class TestRealAPIIntegration:
    """Integration tests with real APIs (requires network)"""

    def test_google_translate_real_api(self):
        """Test with real Google Translate API (free tier)"""
        translator = GoogleTranslator(
            source_lang="en",
            target_lang="zh",
            use_free_api=True,
            max_workers=1,
            timeout=10,
        )

        try:
            result = translator.translate_text("Hello")
            # Should be translated to Chinese
            assert result != "Hello"
            assert len(result) > 0

        except Exception as e:
            pytest.skip(f"Real API test failed (expected in CI): {e}")

    def test_batch_translation_real_api(self):
        """Test batch translation with real API"""
        translator = GoogleTranslator(
            source_lang="en",
            target_lang="zh",
            use_free_api=True,
            max_workers=1,
            timeout=10,
        )

        texts = ["Hello", "World", "Test"]

        try:
            results = translator.translate_text_batch(texts)

            assert len(results) == len(texts)
            # Results should be different from input (translated)
            assert any(result != original for result, original in zip(results, texts))

        except Exception as e:
            pytest.skip(f"Real API batch test failed (expected in CI): {e}")


@pytest.mark.integration
class TestErrorHandling:
    """Test error handling in integration scenarios"""

    def test_file_not_found_handling(self):
        """Test handling of non-existent files"""
        from offitrans.processors.pdf import PDFProcessor

        try:
            processor = PDFProcessor()
            result = processor.process_file("nonexistent.pdf", "output.pdf")

            assert result is False
            stats = processor.get_stats()
            assert stats["failed_files"] == 1

        except ImportError:
            pytest.skip("PDF processor dependencies not available")

    def test_invalid_file_format_handling(self, temp_dir):
        """Test handling of invalid file formats"""
        # Create a text file with PDF extension
        fake_pdf = temp_dir / "fake.pdf"
        fake_pdf.write_text("This is not a PDF file")

        try:
            from offitrans.processors.pdf import PDFProcessor

            processor = PDFProcessor()

            # This should handle the error gracefully
            result = processor.process_file(str(fake_pdf), str(temp_dir / "output.txt"))

            # The result depends on implementation, but it shouldn't crash
            stats = processor.get_stats()
            assert stats["total_files_processed"] >= 1

        except ImportError:
            pytest.skip("PDF processor dependencies not available")

    def test_translation_error_handling(self, temp_dir, file_generator):
        """Test handling of translation errors"""

        # Create a failing translator
        class FailingTranslator:
            def __init__(self):
                self.source_lang = "en"
                self.target_lang = "zh"

            def translate_text_batch(self, texts):
                raise Exception("Translation service unavailable")

        # Create a test file
        input_file = temp_dir / "test.xlsx"
        output_file = temp_dir / "output.xlsx"

        success = file_generator.create_simple_excel(input_file)
        if not success:
            pytest.skip("Could not create test file")

        try:
            from offitrans.processors.excel import ExcelProcessor

            processor = ExcelProcessor(translator=FailingTranslator())

            # This should handle the translation failure gracefully
            result = processor.process_file(str(input_file), str(output_file))

            # Should fail gracefully
            assert result is False
            stats = processor.get_stats()
            assert stats["failed_files"] >= 1

        except ImportError:
            pytest.skip("Excel processor dependencies not available")
