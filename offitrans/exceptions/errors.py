"""
Custom exception classes for Offitrans

This module defines all custom exceptions used throughout the library.
"""


class OffitransError(Exception):
    """
    Base exception class for all Offitrans errors.

    All other custom exceptions in the library inherit from this class.
    """

    def __init__(self, message: str, details: str = None):
        """
        Initialize the exception.

        Args:
            message: Primary error message
            details: Additional error details (optional)
        """
        super().__init__(message)
        self.message = message
        self.details = details

    def __str__(self) -> str:
        """Return string representation of the error."""
        if self.details:
            return f"{self.message}: {self.details}"
        return self.message


class TranslationError(OffitransError):
    """
    Exception raised when translation operations fail.

    This includes API failures, network errors, and translation service issues.
    """

    def __init__(self, message: str, details: str = None, text: str = None):
        """
        Initialize the translation error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            text: The text that failed to translate (optional)
        """
        super().__init__(message, details)
        self.text = text


class ProcessorError(OffitransError):
    """
    Exception raised when file processing operations fail.

    This includes errors in Excel, Word, PDF, or PowerPoint processing.
    """

    def __init__(self, message: str, details: str = None, file_path: str = None):
        """
        Initialize the processor error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            file_path: Path to the file that caused the error (optional)
        """
        super().__init__(message, details)
        self.file_path = file_path


class ConfigError(OffitransError):
    """
    Exception raised when configuration is invalid or missing.

    This includes missing API keys, invalid settings, and configuration validation errors.
    """

    def __init__(self, message: str, details: str = None, config_key: str = None):
        """
        Initialize the configuration error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            config_key: The configuration key that caused the error (optional)
        """
        super().__init__(message, details)
        self.config_key = config_key


class FileError(OffitransError):
    """
    Exception raised when file operations fail.

    This includes file not found, permission errors, and file format issues.
    """

    def __init__(self, message: str, details: str = None, file_path: str = None):
        """
        Initialize the file error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            file_path: Path to the file that caused the error (optional)
        """
        super().__init__(message, details)
        self.file_path = file_path


class APIError(OffitransError):
    """
    Exception raised when API operations fail.

    This includes authentication errors, rate limiting, and service unavailability.
    """

    def __init__(
        self,
        message: str,
        details: str = None,
        status_code: int = None,
        response_body: str = None,
    ):
        """
        Initialize the API error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            status_code: HTTP status code (optional)
            response_body: Response body from the API (optional)
        """
        super().__init__(message, details)
        self.status_code = status_code
        self.response_body = response_body


class CacheError(OffitransError):
    """
    Exception raised when cache operations fail.

    This includes cache file corruption, permission errors, and disk space issues.
    """

    def __init__(self, message: str, details: str = None, cache_file: str = None):
        """
        Initialize the cache error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            cache_file: Path to the cache file that caused the error (optional)
        """
        super().__init__(message, details)
        self.cache_file = cache_file


# Specialized processor errors for different file types


class ExcelProcessorError(ProcessorError):
    """Exception specific to Excel file processing."""

    def __init__(
        self,
        message: str,
        details: str = None,
        file_path: str = None,
        sheet_name: str = None,
        cell_address: str = None,
    ):
        """
        Initialize Excel processor error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            file_path: Path to the Excel file (optional)
            sheet_name: Name of the worksheet (optional)
            cell_address: Cell address where error occurred (optional)
        """
        super().__init__(message, details, file_path)
        self.sheet_name = sheet_name
        self.cell_address = cell_address


class WordProcessorError(ProcessorError):
    """Exception specific to Word document processing."""

    def __init__(
        self,
        message: str,
        details: str = None,
        file_path: str = None,
        paragraph_index: int = None,
        run_index: int = None,
    ):
        """
        Initialize Word processor error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            file_path: Path to the Word document (optional)
            paragraph_index: Index of the paragraph where error occurred (optional)
            run_index: Index of the run where error occurred (optional)
        """
        super().__init__(message, details, file_path)
        self.paragraph_index = paragraph_index
        self.run_index = run_index


class PDFProcessorError(ProcessorError):
    """Exception specific to PDF file processing."""

    def __init__(
        self,
        message: str,
        details: str = None,
        file_path: str = None,
        page_number: int = None,
    ):
        """
        Initialize PDF processor error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            file_path: Path to the PDF file (optional)
            page_number: Page number where error occurred (optional)
        """
        super().__init__(message, details, file_path)
        self.page_number = page_number


class PowerPointProcessorError(ProcessorError):
    """Exception specific to PowerPoint file processing."""

    def __init__(
        self,
        message: str,
        details: str = None,
        file_path: str = None,
        slide_index: int = None,
        shape_index: int = None,
    ):
        """
        Initialize PowerPoint processor error.

        Args:
            message: Primary error message
            details: Additional error details (optional)
            file_path: Path to the PowerPoint file (optional)
            slide_index: Index of the slide where error occurred (optional)
            shape_index: Index of the shape where error occurred (optional)
        """
        super().__init__(message, details, file_path)
        self.slide_index = slide_index
        self.shape_index = shape_index
