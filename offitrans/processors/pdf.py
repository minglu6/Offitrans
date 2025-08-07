"""
PDF file processor for Offitrans

This module provides functionality to translate PDF files.
Note: PDF translation is complex and may have limitations.
"""

import logging
from typing import List, Dict, Any
from pathlib import Path

try:
    import PyPDF2

    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

from .base import BaseProcessor
from ..exceptions.errors import PDFProcessorError

logger = logging.getLogger(__name__)


class PDFProcessor(BaseProcessor):
    """
    PDF file processor that handles text extraction and translation.

    Note: This is a basic implementation. PDF translation is complex
    due to the nature of PDF format and layout preservation challenges.
    """

    def __init__(self, **kwargs):
        """
        Initialize PDF processor.

        Args:
            **kwargs: Additional arguments passed to BaseProcessor
        """
        if not PYPDF2_AVAILABLE:
            raise PDFProcessorError(
                "PyPDF2 library is required for PDF processing",
                details="Install with: pip install PyPDF2",
            )

        super().__init__(**kwargs)

    def supports_file_type(self, file_path: str) -> bool:
        """
        Check if file type is supported.

        Args:
            file_path: Path to the file

        Returns:
            True if file type is supported
        """
        supported_extensions = {".pdf"}
        return Path(file_path).suffix.lower() in supported_extensions

    def extract_text(self, file_path: str) -> List[Dict[str, Any]]:
        """
        Extract text content from PDF file.

        Args:
            file_path: Path to the PDF file

        Returns:
            List of dictionaries containing text and metadata
        """
        text_data = []

        try:
            with open(file_path, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                logger.info(f"Successfully opened PDF file: {file_path}")
                logger.info(f"PDF has {len(reader.pages)} pages")

                for page_num, page in enumerate(reader.pages):
                    try:
                        page_text = page.extract_text()

                        if page_text.strip():
                            # Split page text into paragraphs
                            paragraphs = [
                                p.strip() for p in page_text.split("\n\n") if p.strip()
                            ]

                            for para_idx, paragraph in enumerate(paragraphs):
                                if paragraph:
                                    text_data.append(
                                        {
                                            "text": paragraph,
                                            "page_number": page_num + 1,
                                            "paragraph_index": para_idx,
                                            "type": "paragraph",
                                        }
                                    )

                                    logger.debug(
                                        f"Extracted text from page {page_num + 1}, paragraph {para_idx}: '{paragraph[:50]}...'"
                                    )

                    except Exception as e:
                        logger.error(
                            f"Error extracting text from page {page_num + 1}: {e}"
                        )
                        continue

            logger.info(f"Total extracted {len(text_data)} text elements from PDF")
            return text_data

        except Exception as e:
            raise PDFProcessorError(
                f"Failed to extract text from PDF file",
                details=str(e),
                file_path=file_path,
            ) from e

    def translate_and_save(
        self, file_path: str, output_path: str, target_language: str = "en"
    ) -> bool:
        """
        Translate PDF file and save to output path.

        Note: This creates a new PDF with translated text.
        The original layout may not be preserved perfectly.

        Args:
            file_path: Path to input PDF file
            output_path: Path for output PDF file
            target_language: Target language code

        Returns:
            True if successful, False otherwise
        """
        try:
            # Step 1: Extract text and metadata
            logger.info("Step 1: Extracting text from PDF file...")
            text_data = self.extract_text(file_path)

            if not text_data:
                logger.warning("No translatable text found in PDF file")
                return False

            # Step 2: Preprocess and translate texts
            logger.info("Step 2: Translating texts...")
            original_texts = [item["text"] for item in text_data]
            unique_texts, metadata = self.preprocess_texts(original_texts)
            translated_unique = self.translate_texts(unique_texts, target_language)
            translated_texts = self.postprocess_translations(
                original_texts, translated_unique, metadata
            )

            # Step 3: Save translated text to new PDF or text file
            logger.info("Step 3: Saving translated content...")
            success = self._save_translated_content(
                text_data, translated_texts, output_path
            )

            if success:
                logger.info(f"Successfully translated PDF content: {output_path}")
                return True
            else:
                logger.error("Failed to save translated PDF content")
                return False

        except Exception as e:
            logger.error(f"Error translating PDF file: {e}")
            return False

    def _save_translated_content(
        self,
        text_data: List[Dict[str, Any]],
        translated_texts: List[str],
        output_path: str,
    ) -> bool:
        """
        Save translated content to output file.

        Currently saves as text file due to complexity of PDF generation.
        Future versions could use reportlab or similar libraries.

        Args:
            text_data: Original text data with metadata
            translated_texts: List of translated texts
            output_path: Output file path

        Returns:
            True if successful, False otherwise
        """
        try:
            # Change extension to .txt if it's .pdf
            output_path = Path(output_path)
            if output_path.suffix.lower() == ".pdf":
                output_path = output_path.with_suffix(".txt")
                logger.info(f"Saving as text file: {output_path}")

            with open(output_path, "w", encoding="utf-8") as f:
                f.write("Translated PDF Content\n")
                f.write("=" * 50 + "\n\n")

                current_page = None
                for item, translated_text in zip(text_data, translated_texts):
                    page_num = item.get("page_number", 1)

                    # Add page header if page changed
                    if current_page != page_num:
                        if current_page is not None:
                            f.write("\n" + "-" * 30 + "\n\n")
                        f.write(f"Page {page_num}\n")
                        f.write("-" * 10 + "\n\n")
                        current_page = page_num

                    f.write(translated_text + "\n\n")

            logger.info(f"Successfully saved translated content to: {output_path}")
            return True

        except Exception as e:
            logger.error(f"Error saving translated content: {e}")
            return False


# Function for simple PDF translation (backward compatibility)
def translate_pdf(
    input_path: str, output_path: str, target_language: str = "en"
) -> bool:
    """
    Simple function to translate a PDF file.

    Args:
        input_path: Path to input PDF file
        output_path: Path to output file
        target_language: Target language code

    Returns:
        True if successful, False otherwise
    """
    try:
        processor = PDFProcessor()
        return processor.process_file(input_path, output_path, target_language)
    except Exception as e:
        logger.error(f"Error in translate_pdf: {e}")
        return False
