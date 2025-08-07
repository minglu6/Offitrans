"""
File processors for Offitrans

This module contains processors for different Office file formats.
"""

from .base import BaseProcessor
from .excel import ExcelProcessor
from .word import WordProcessor
from .pdf import PDFProcessor
from .powerpoint import PowerPointProcessor

__all__ = [
    "BaseProcessor",
    "ExcelProcessor",
    "WordProcessor",
    "PDFProcessor",
    "PowerPointProcessor",
]

# Available processor types
AVAILABLE_PROCESSORS = {
    "excel": ExcelProcessor,
    "xlsx": ExcelProcessor,
    "xls": ExcelProcessor,
    "word": WordProcessor,
    "docx": WordProcessor,
    "doc": WordProcessor,
    "pdf": PDFProcessor,
    "powerpoint": PowerPointProcessor,
    "pptx": PowerPointProcessor,
    "ppt": PowerPointProcessor,
}


def get_processor(file_type: str, **kwargs):
    """
    Get a processor instance by file type.

    Args:
        file_type: Type of file processor (e.g., 'excel', 'word', 'pdf', 'powerpoint')
        **kwargs: Arguments to pass to processor constructor

    Returns:
        Processor instance

    Raises:
        ValueError: If file type is not supported
    """
    file_type = file_type.lower()

    if file_type not in AVAILABLE_PROCESSORS:
        available = ", ".join(set(AVAILABLE_PROCESSORS.keys()))
        raise ValueError(f"Unsupported file type: {file_type}. Available: {available}")

    processor_class = AVAILABLE_PROCESSORS[file_type]
    return processor_class(**kwargs)


def get_processor_by_extension(file_path: str, **kwargs):
    """
    Get a processor instance by file extension.

    Args:
        file_path: Path to the file
        **kwargs: Arguments to pass to processor constructor

    Returns:
        Processor instance

    Raises:
        ValueError: If file extension is not supported
    """
    from pathlib import Path

    extension = Path(file_path).suffix.lower().lstrip(".")
    return get_processor(extension, **kwargs)
