"""Core utilities package."""
from .faculty_manager import FacultyManager
from .excel_styler import (
    style_merged_range,
    apply_table_header,
    apply_data_row,
    autofit_column_widths,
    FONT_TITLE,
    FONT_DEPT,
    FONT_TABLE_HEADER,
    FONT_DATA
)

__all__ = [
    "FacultyManager",
    "style_merged_range",
    "apply_table_header",
    "apply_data_row",
    "autofit_column_widths",
    "FONT_TITLE",
    "FONT_DEPT",
    "FONT_TABLE_HEADER",
    "FONT_DATA"
]
