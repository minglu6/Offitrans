#!/usr/bin/env python3
"""Main CLI entry point for Offitrans."""

import argparse
import sys
from pathlib import Path
from typing import Optional

from ..processors.excel import ExcelProcessor
from ..processors.word import WordProcessor
from ..processors.pdf import PDFProcessor
from ..processors.powerpoint import PowerPointProcessor
from ..translators.google import GoogleTranslator
from ..core.config import Config


def create_parser() -> argparse.ArgumentParser:
    """Create the command line argument parser."""
    parser = argparse.ArgumentParser(
        prog="offitrans",
        description="Translate office documents (Excel, Word, PDF, PowerPoint)",
    )

    parser.add_argument(
        "input_file", type=str, help="Path to the input file to translate"
    )

    parser.add_argument(
        "-o",
        "--output",
        type=str,
        help="Output file path (default: adds '_translated' suffix to input filename)",
    )

    parser.add_argument(
        "-s",
        "--source",
        type=str,
        default="auto",
        help="Source language (default: auto-detect)",
    )

    parser.add_argument(
        "-t",
        "--target",
        type=str,
        required=True,
        help="Target language code (e.g., 'en', 'zh', 'fr')",
    )

    parser.add_argument(
        "--translator",
        type=str,
        default="google",
        choices=["google"],
        help="Translation service to use (default: google)",
    )

    parser.add_argument(
        "--api-key", type=str, help="API key for translation service (if required)"
    )

    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Enable verbose output"
    )

    parser.add_argument("--version", action="version", version="%(prog)s 0.2.0")

    return parser


def get_processor(file_path: Path):
    """Get appropriate processor for the file type."""
    suffix = file_path.suffix.lower()

    if suffix in [".xlsx", ".xls"]:
        return ExcelProcessor()
    elif suffix in [".docx", ".doc"]:
        return WordProcessor()
    elif suffix == ".pdf":
        return PDFProcessor()
    elif suffix in [".pptx", ".ppt"]:
        return PowerPointProcessor()
    else:
        raise ValueError(f"Unsupported file type: {suffix}")


def get_translator(translator_name: str, api_key: Optional[str] = None):
    """Get appropriate translator instance."""
    if translator_name == "google":
        return GoogleTranslator(api_key=api_key)
    else:
        raise ValueError(f"Unsupported translator: {translator_name}")


def main() -> int:
    """Main CLI entry point."""
    parser = create_parser()
    args = parser.parse_args()

    try:
        # Validate input file
        input_path = Path(args.input_file)
        if not input_path.exists():
            print(f"Error: Input file '{input_path}' does not exist.", file=sys.stderr)
            return 1

        if not input_path.is_file():
            print(f"Error: '{input_path}' is not a file.", file=sys.stderr)
            return 1

        # Determine output path
        if args.output:
            output_path = Path(args.output)
        else:
            output_path = (
                input_path.parent / f"{input_path.stem}_translated{input_path.suffix}"
            )

        if args.verbose:
            print(f"Input file: {input_path}")
            print(f"Output file: {output_path}")
            print(f"Source language: {args.source}")
            print(f"Target language: {args.target}")
            print(f"Translator: {args.translator}")

        # Get processor and translator
        processor = get_processor(input_path)
        translator = get_translator(args.translator, args.api_key)

        # Perform translation
        if args.verbose:
            print("Starting translation...")

        processor.translate_file(
            input_path=str(input_path),
            output_path=str(output_path),
            translator=translator,
            source_lang=args.source,
            target_lang=args.target,
        )

        print(f"Translation completed successfully: {output_path}")
        return 0

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback

            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
