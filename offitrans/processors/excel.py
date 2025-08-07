"""
Excel file processor for Offitrans

This module provides functionality to translate Excel files while preserving
formatting, images, and layout.
"""

import os
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.styles.colors import Color
    from openpyxl.cell.text import InlineFont
    from openpyxl.cell.rich_text import TextBlock, CellRichText
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from PIL import Image as PILImage

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from .base import BaseProcessor
from ..exceptions.errors import ExcelProcessorError

logger = logging.getLogger(__name__)


class ExcelProcessor(BaseProcessor):
    """
    Excel file processor that handles translation while preserving formatting.

    This processor can handle complex Excel files with:
    - Rich text formatting
    - Merged cells
    - Images and charts
    - Multiple worksheets
    - Various cell formats and styles
    """

    def __init__(self, **kwargs):
        """
        Initialize Excel processor.

        Args:
            **kwargs: Additional arguments passed to BaseProcessor
        """
        if not OPENPYXL_AVAILABLE:
            raise ExcelProcessorError(
                "openpyxl library is required for Excel processing",
                details="Install with: pip install openpyxl",
            )

        super().__init__(**kwargs)

        # Excel-specific settings
        self.smart_column_width = getattr(
            self.config.processor, "smart_column_width", True
        )

        # Image data storage
        self.image_data: Dict[str, List[Dict[str, Any]]] = {}

    def supports_file_type(self, file_path: str) -> bool:
        """
        Check if file type is supported.

        Args:
            file_path: Path to the file

        Returns:
            True if file type is supported
        """
        supported_extensions = {".xlsx", ".xls", ".xlsm"}
        return Path(file_path).suffix.lower() in supported_extensions

    def extract_images_info(self, workbook) -> Dict[str, List[Dict[str, Any]]]:
        """
        Extract image information from Excel workbook.

        Args:
            workbook: openpyxl workbook object

        Returns:
            Dictionary mapping sheet names to image info lists
        """
        images_info = {}

        try:
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                sheet_images = []

                # Check for images in the sheet
                if hasattr(sheet, "_images") and sheet._images:
                    logger.info(
                        f"Found {len(sheet._images)} images in sheet '{sheet_name}'"
                    )

                    for img in sheet._images:
                        img_info = {
                            "image_object": img,
                            "anchor_type": type(img.anchor).__name__,
                        }

                        # Extract anchor information
                        if isinstance(img.anchor, TwoCellAnchor):
                            img_info["anchor_info"] = {
                                "type": "two_cell",
                                "from_col": img.anchor._from.col,
                                "from_col_off": img.anchor._from.colOff,
                                "from_row": img.anchor._from.row,
                                "from_row_off": img.anchor._from.rowOff,
                                "to_col": img.anchor.to.col,
                                "to_col_off": img.anchor.to.colOff,
                                "to_row": img.anchor.to.row,
                                "to_row_off": img.anchor.to.rowOff,
                            }
                        elif isinstance(img.anchor, OneCellAnchor):
                            img_info["anchor_info"] = {
                                "type": "one_cell",
                                "from_col": img.anchor._from.col,
                                "from_col_off": img.anchor._from.colOff,
                                "from_row": img.anchor._from.row,
                                "from_row_off": img.anchor._from.rowOff,
                                "width": img.anchor.ext.cx,
                                "height": img.anchor.ext.cy,
                            }

                        sheet_images.append(img_info)

                images_info[sheet_name] = sheet_images

        except Exception as e:
            logger.error(f"Error extracting image info: {e}")

        return images_info

    def restore_images_info(
        self, workbook, images_info: Dict[str, List[Dict[str, Any]]]
    ) -> None:
        """
        Restore image information to Excel workbook.

        Args:
            workbook: openpyxl workbook object
            images_info: Dictionary with image information
        """
        try:
            for sheet_name, sheet_images in images_info.items():
                if not sheet_images:
                    continue

                sheet = workbook[sheet_name]

                # Clear existing images
                if hasattr(sheet, "_images"):
                    sheet._images.clear()
                else:
                    sheet._images = []

                # Restore images
                for img_info in sheet_images:
                    try:
                        img_obj = img_info["image_object"]

                        # Use the original image object if possible
                        if hasattr(img_obj, "anchor") and img_obj.anchor:
                            sheet.add_image(img_obj)
                            logger.debug(
                                f"Successfully restored image in sheet {sheet_name}"
                            )
                        else:
                            logger.warning(
                                f"Could not restore image in sheet {sheet_name}"
                            )

                    except Exception as e:
                        logger.error(f"Error restoring image: {e}")
                        continue

        except Exception as e:
            logger.error(f"Error restoring images: {e}")

    def extract_text(self, file_path: str) -> List[Dict[str, Any]]:
        """
        Extract text content from Excel file.

        Args:
            file_path: Path to the Excel file

        Returns:
            List of dictionaries containing text and metadata
        """
        text_data = []

        try:
            workbook = load_workbook(file_path, data_only=False)
            logger.info(f"Successfully opened Excel file: {file_path}")

            # Extract image information if image protection is enabled
            if self.image_protection:
                logger.info("Extracting image information...")
                self.image_data = self.extract_images_info(workbook)

            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                logger.info(f"Processing worksheet: {sheet_name}")

                # Iterate through all cells
                for row in sheet.iter_rows():
                    for cell in row:
                        if (
                            cell.value
                            and isinstance(cell.value, str)
                            and cell.value.strip()
                        ):
                            # Skip formula cells (starting with =)
                            if not cell.value.startswith("="):
                                # Extract cell format information
                                format_info = self._extract_cell_format(cell)

                                # Check for rich text formatting
                                rich_text_info = self._extract_rich_text_format(cell)

                                text_data.append(
                                    {
                                        "text": cell.value,
                                        "sheet_name": sheet_name,
                                        "row": cell.row,
                                        "column": cell.column,
                                        "cell_coordinate": cell.coordinate,
                                        "format_info": format_info,
                                        "rich_text_info": rich_text_info,
                                    }
                                )

                                logger.debug(
                                    f"Extracted text from {sheet_name}!{cell.coordinate}: '{cell.value[:50]}...'"
                                )

            workbook.close()
            logger.info(f"Total extracted {len(text_data)} text cells")
            return text_data

        except Exception as e:
            raise ExcelProcessorError(
                f"Failed to extract text from Excel file",
                details=str(e),
                file_path=file_path,
            ) from e

    def translate_and_save(
        self, file_path: str, output_path: str, target_language: str = "en"
    ) -> bool:
        """
        Translate Excel file and save to output path.

        Args:
            file_path: Path to input Excel file
            output_path: Path for output Excel file
            target_language: Target language code

        Returns:
            True if successful, False otherwise
        """
        try:
            # Step 1: Extract text and metadata
            logger.info("Step 1: Extracting text from Excel file...")
            text_data = self.extract_text(file_path)

            if not text_data:
                logger.warning("No translatable text found in Excel file")
                return False

            # Step 2: Preprocess and translate texts
            logger.info("Step 2: Translating texts...")
            original_texts = [item["text"] for item in text_data]
            unique_texts, metadata = self.preprocess_texts(original_texts)
            translated_unique = self.translate_texts(unique_texts, target_language)
            translated_texts = self.postprocess_translations(
                original_texts, translated_unique, metadata
            )

            # Step 3: Apply translations to Excel file
            logger.info("Step 3: Applying translations to Excel file...")
            success = self._replace_text_with_format_and_images(
                file_path, output_path, text_data, translated_texts, target_language
            )

            if success:
                logger.info(f"Successfully translated Excel file: {output_path}")
                return True
            else:
                logger.error("Failed to apply translations to Excel file")
                return False

        except Exception as e:
            logger.error(f"Error translating Excel file: {e}")
            return False

    def _replace_text_with_format_and_images(
        self,
        excel_path: str,
        output_path: str,
        text_data: List[Dict[str, Any]],
        translated_texts: List[str],
        target_language: str = "en",
    ) -> bool:
        """
        Replace text in Excel file while preserving formatting and images.

        Args:
            excel_path: Input Excel file path
            output_path: Output Excel file path
            text_data: Original text data with metadata
            translated_texts: List of translated texts
            target_language: Target language code

        Returns:
            True if successful, False otherwise
        """
        try:
            workbook = load_workbook(excel_path, data_only=False)

            # Replace text in cells
            for item, translated_text in zip(text_data, translated_texts):
                sheet_name = item["sheet_name"]
                row = item["row"]
                column = item["column"]
                format_info = item["format_info"]

                # Get worksheet and cell
                sheet = workbook[sheet_name]
                cell = sheet.cell(row=row, column=column)

                # Replace text
                cell.value = translated_text

                # Apply formatting
                self._apply_cell_format(cell, format_info, target_language)

                # Apply rich text formatting if available
                rich_text_info = item.get("rich_text_info")
                if rich_text_info and rich_text_info.get("has_rich_text"):
                    self._apply_rich_text_format(
                        cell,
                        item["text"],
                        translated_text,
                        rich_text_info,
                        target_language,
                    )

                logger.debug(f"Applied translation to {sheet_name}!{cell.coordinate}")

            # Restore images if image protection is enabled
            if self.image_protection and self.image_data:
                logger.info("Restoring image information...")
                self.restore_images_info(workbook, self.image_data)

            # Apply smart column width adjustment if enabled
            if self.smart_column_width:
                logger.info("Adjusting column widths...")
                self._smart_adjust_column_width(workbook)

            # Save the workbook
            workbook.save(output_path)
            workbook.close()

            logger.info(f"Successfully saved translated Excel file: {output_path}")
            return True

        except Exception as e:
            logger.error(f"Error replacing text in Excel file: {e}")
            return False

    def _extract_cell_format(self, cell) -> Dict[str, Any]:
        """
        Extract formatting information from a cell.

        Args:
            cell: openpyxl cell object

        Returns:
            Dictionary containing format information
        """
        format_info = {}

        try:
            # Font information
            if cell.font:
                format_info["font_name"] = cell.font.name
                format_info["font_size"] = cell.font.size
                format_info["font_bold"] = cell.font.bold
                format_info["font_italic"] = cell.font.italic
                format_info["font_underline"] = cell.font.underline
                format_info["font_strike"] = cell.font.strike

                # Color information
                if cell.font.color:
                    format_info["font_color"] = cell.font.color
                    if hasattr(cell.font.color, "rgb") and cell.font.color.rgb:
                        format_info["font_color_rgb"] = cell.font.color.rgb
                    elif (
                        hasattr(cell.font.color, "indexed")
                        and cell.font.color.indexed is not None
                    ):
                        format_info["font_color_indexed"] = cell.font.color.indexed
                    elif (
                        hasattr(cell.font.color, "theme")
                        and cell.font.color.theme is not None
                    ):
                        format_info["font_color_theme"] = cell.font.color.theme
                        if (
                            hasattr(cell.font.color, "tint")
                            and cell.font.color.tint is not None
                        ):
                            format_info["font_color_tint"] = cell.font.color.tint

            # Fill information
            if cell.fill and hasattr(cell.fill, "start_color"):
                format_info["fill_color"] = cell.fill.start_color
                format_info["fill_type"] = cell.fill.fill_type
                format_info["fill_object"] = cell.fill

            # Alignment information
            if cell.alignment:
                format_info["horizontal"] = cell.alignment.horizontal
                format_info["vertical"] = cell.alignment.vertical
                format_info["wrap_text"] = cell.alignment.wrap_text
                format_info["shrink_to_fit"] = cell.alignment.shrink_to_fit

            # Border information
            if cell.border:
                format_info["has_border"] = True
                format_info["border"] = cell.border

            # Number format
            if cell.number_format:
                format_info["number_format"] = cell.number_format

        except Exception as e:
            logger.error(f"Error extracting cell format: {e}")

        return format_info

    def _extract_rich_text_format(self, cell) -> Optional[Dict[str, Any]]:
        """
        Extract rich text formatting information from a cell.

        Args:
            cell: openpyxl cell object

        Returns:
            Rich text format information or None
        """
        try:
            # Check for rich text in cell value
            if hasattr(cell, "_value") and isinstance(cell._value, CellRichText):
                return self._parse_rich_text_object(cell._value, cell.coordinate)
            elif isinstance(cell.value, CellRichText):
                return self._parse_rich_text_object(cell.value, cell.coordinate)

            return None

        except Exception as e:
            logger.error(f"Error extracting rich text format: {e}")
            return None

    def _parse_rich_text_object(
        self, rich_text: CellRichText, coordinate: str
    ) -> Dict[str, Any]:
        """
        Parse rich text object and extract formatting information.

        Args:
            rich_text: CellRichText object
            coordinate: Cell coordinate

        Returns:
            Rich text information dictionary
        """
        rich_info = {"has_rich_text": True, "segments": []}

        logger.debug(f"Found rich text format in {coordinate}")

        try:
            for i, item in enumerate(rich_text):
                if isinstance(item, TextBlock):
                    segment_info = {"text": item.text, "font": None, "segment_index": i}

                    # Extract font information
                    if item.font:
                        font_info = {
                            "name": getattr(item.font, "rFont", None),
                            "size": getattr(item.font, "sz", None),
                            "bold": getattr(item.font, "b", None),
                            "italic": getattr(item.font, "i", None),
                            "underline": getattr(item.font, "u", None),
                            "color": getattr(item.font, "color", None),
                        }
                        segment_info["font"] = font_info

                    rich_info["segments"].append(segment_info)

                elif isinstance(item, str):
                    # Plain text segment
                    rich_info["segments"].append(
                        {"text": item, "font": None, "segment_index": i}
                    )

        except Exception as e:
            logger.error(f"Error parsing rich text object: {e}")

        return rich_info

    def _apply_cell_format(
        self, cell, format_info: Dict[str, Any], target_language: str = "en"
    ) -> None:
        """
        Apply formatting to a cell.

        Args:
            cell: openpyxl cell object
            format_info: Format information dictionary
            target_language: Target language code
        """
        try:
            if not format_info:
                return

            # Apply font formatting
            font_kwargs = {}

            # Font name (with language-specific adjustments)
            if target_language == "th" and format_info.get("font_name"):
                # Use Thai-compatible font
                thai_fonts = ["TH SarabunPSK", "Tahoma", "Arial Unicode MS"]
                original_font = format_info["font_name"]
                if not any(
                    thai_font.lower() in original_font.lower()
                    for thai_font in thai_fonts
                ):
                    font_kwargs["name"] = "TH SarabunPSK"
                else:
                    font_kwargs["name"] = original_font
            elif format_info.get("font_name"):
                font_kwargs["name"] = format_info["font_name"]

            # Font size (with adjustment)
            if format_info.get("font_size"):
                original_size = format_info["font_size"]
                adjusted_size = max(6, int(original_size * self.font_size_adjustment))
                font_kwargs["size"] = adjusted_size

            # Other font properties
            for prop in ["font_bold", "font_italic", "font_underline", "font_strike"]:
                if format_info.get(prop) is not None:
                    font_kwargs[prop.replace("font_", "")] = format_info[prop]

            # Font color - create a new Color object to avoid StyleProxy issues
            if format_info.get("font_color"):
                try:
                    color_obj = format_info["font_color"]
                    if hasattr(color_obj, "rgb") and color_obj.rgb:
                        font_kwargs["color"] = Color(rgb=color_obj.rgb)
                    elif (
                        hasattr(color_obj, "indexed") and color_obj.indexed is not None
                    ):
                        font_kwargs["color"] = Color(indexed=color_obj.indexed)
                    elif hasattr(color_obj, "theme") and color_obj.theme is not None:
                        tint = getattr(color_obj, "tint", None)
                        font_kwargs["color"] = Color(theme=color_obj.theme, tint=tint)
                except Exception as e:
                    logger.debug(f"Could not apply font color: {e}")

            if font_kwargs:
                cell.font = Font(**font_kwargs)

            # Apply fill formatting - create a new PatternFill to avoid StyleProxy issues
            if format_info.get("fill_object"):
                try:
                    fill_obj = format_info["fill_object"]
                    if hasattr(fill_obj, "fill_type") and hasattr(
                        fill_obj, "start_color"
                    ):
                        start_color = fill_obj.start_color
                        if hasattr(start_color, "rgb") and start_color.rgb:
                            cell.fill = PatternFill(
                                fill_type=fill_obj.fill_type,
                                start_color=Color(rgb=start_color.rgb),
                            )
                        elif (
                            hasattr(start_color, "indexed")
                            and start_color.indexed is not None
                        ):
                            cell.fill = PatternFill(
                                fill_type=fill_obj.fill_type,
                                start_color=Color(indexed=start_color.indexed),
                            )
                        elif (
                            hasattr(start_color, "theme")
                            and start_color.theme is not None
                        ):
                            tint = getattr(start_color, "tint", None)
                            cell.fill = PatternFill(
                                fill_type=fill_obj.fill_type,
                                start_color=Color(theme=start_color.theme, tint=tint),
                            )
                except Exception as e:
                    logger.debug(f"Could not apply fill formatting: {e}")

            # Apply alignment
            alignment_kwargs = {}
            for prop in ["horizontal", "vertical", "wrap_text", "shrink_to_fit"]:
                if format_info.get(prop) is not None:
                    alignment_kwargs[prop] = format_info[prop]

            if alignment_kwargs:
                cell.alignment = Alignment(**alignment_kwargs)

            # Apply border - create new Border object to avoid StyleProxy issues
            if format_info.get("border"):
                try:
                    # For now, skip border application to avoid StyleProxy issues
                    # A more complete implementation would recreate the border object
                    pass
                except Exception as e:
                    logger.debug(f"Could not apply border: {e}")

            # Apply number format
            if format_info.get("number_format"):
                cell.number_format = format_info["number_format"]

        except Exception as e:
            logger.error(f"Error applying cell format: {e}")

    def _apply_rich_text_format(
        self,
        cell,
        original_text: str,
        translated_text: str,
        rich_text_info: Dict[str, Any],
        target_language: str = "en",
    ) -> None:
        """
        Apply rich text formatting to translated text.

        Args:
            cell: openpyxl cell object
            original_text: Original text
            translated_text: Translated text
            rich_text_info: Rich text formatting information
            target_language: Target language code
        """
        if not rich_text_info or not rich_text_info.get("has_rich_text"):
            return

        try:
            segments = rich_text_info.get("segments", [])
            if not segments:
                return

            # Create new rich text parts
            rich_text_parts = []

            # For simplicity, apply the first segment's format to the entire translated text
            # More sophisticated mapping could be implemented based on text length ratios
            if segments:
                first_segment = segments[0]
                if first_segment.get("font"):
                    # Create inline font
                    font_info = first_segment["font"]
                    inline_font = self._create_inline_font(font_info, target_language)
                    rich_text_parts.append(TextBlock(inline_font, translated_text))
                else:
                    rich_text_parts.append(translated_text)

            # Apply rich text to cell
            if rich_text_parts:
                cell._value = CellRichText(rich_text_parts)

        except Exception as e:
            logger.error(f"Error applying rich text format: {e}")
            # Fall back to plain text
            cell.value = translated_text

    def _create_inline_font(
        self, font_info: Dict[str, Any], target_language: str = "en"
    ) -> InlineFont:
        """
        Create an InlineFont object from font information.

        Args:
            font_info: Font information dictionary
            target_language: Target language code

        Returns:
            InlineFont object
        """
        font_kwargs = {}

        if font_info.get("name"):
            font_kwargs["rFont"] = font_info["name"]
        elif target_language == "th":
            font_kwargs["rFont"] = "TH SarabunPSK"

        if font_info.get("size"):
            font_kwargs["sz"] = font_info["size"]
        if font_info.get("bold"):
            font_kwargs["b"] = font_info["bold"]
        if font_info.get("italic"):
            font_kwargs["i"] = font_info["italic"]
        if font_info.get("underline"):
            font_kwargs["u"] = font_info["underline"]
        if font_info.get("color"):
            font_kwargs["color"] = font_info["color"]

        return InlineFont(**font_kwargs)

    def _smart_adjust_column_width(self, workbook) -> None:
        """
        Intelligently adjust column widths to fit content.

        Args:
            workbook: openpyxl workbook object
        """
        try:
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Get image-occupied columns if image protection is enabled
                occupied_columns = set()
                if self.image_protection and sheet_name in self.image_data:
                    for img_info in self.image_data[sheet_name]:
                        anchor_info = img_info.get("anchor_info", {})
                        if anchor_info.get("type") == "two_cell":
                            from_col = anchor_info.get("from_col", 0)
                            to_col = anchor_info.get("to_col", 0)
                            for col in range(from_col, to_col + 1):
                                occupied_columns.add(col)

                # Adjust column widths
                for column in sheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    column_index = column[0].column

                    # Check if column has images
                    is_occupied = column_index in occupied_columns

                    for cell in column:
                        try:
                            if cell.value:
                                cell_length = len(str(cell.value))
                                if cell_length > max_length:
                                    max_length = cell_length
                        except Exception:
                            pass

                    # Set column width (conservative for image-occupied columns)
                    if is_occupied:
                        adjusted_width = min(max_length + 1, 30)
                    else:
                        adjusted_width = min(max_length + 2, 50)

                    sheet.column_dimensions[column_letter].width = adjusted_width

        except Exception as e:
            logger.error(f"Error adjusting column widths: {e}")


# Backward compatibility alias
ExcelTranslator = ExcelProcessor
