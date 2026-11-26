"""Write captured balance tables directly to Excel via openpyxl (no COM / clipboard)."""

from copy import copy
from decimal import Decimal, InvalidOperation
from html.parser import HTMLParser
import os
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from extractors.base import BalanceTableCapture


_NAMED_COLORS: Dict[str, str] = {
    "black": "000000",
    "white": "FFFFFF",
    "red": "FF0000",
    "green": "008000",
    "blue": "0000FF",
    "yellow": "FFFF00",
    "cyan": "00FFFF",
    "magenta": "FF00FF",
    "gray": "808080",
    "grey": "808080",
    "lightgray": "D3D3D3",
    "lightgrey": "D3D3D3",
    "darkgray": "A9A9A9",
    "darkgrey": "A9A9A9",
    "orange": "FFA500",
    "silver": "C0C0C0",
    "navy": "000080",
    "teal": "008080",
    "purple": "800080",
    "maroon": "800000",
}


def _parse_css_declarations(style_text: str) -> Dict[str, str]:
    declarations: Dict[str, str] = {}
    if not style_text:
        return declarations
    for part in style_text.split(";"):
        if ":" not in part:
            continue
        key, value = part.split(":", 1)
        key = key.strip().lower()
        value = value.strip()
        if key and value:
            declarations[key] = value
    return declarations


def _parse_css_color(value: Optional[str]) -> Optional[str]:
    """Convert a CSS color string to an uppercase 6-char RGB hex (alpha blended onto white)."""
    if not value:
        return None
    text = value.strip().lower()
    if text in ("none", "transparent", "inherit", "initial"):
        return None
    if text in _NAMED_COLORS:
        return _NAMED_COLORS[text]

    hex_match = re.fullmatch(r"#([0-9a-f]{3}|[0-9a-f]{6})", text)
    if hex_match:
        raw = hex_match.group(1)
        if len(raw) == 3:
            raw = "".join(ch * 2 for ch in raw)
        return raw.upper()

    rgb_match = re.fullmatch(
        r"rgba?\(\s*([0-9.]+)\s*,\s*([0-9.]+)\s*,\s*([0-9.]+)(?:\s*,\s*([0-9.]+))?\s*\)",
        text,
    )
    if rgb_match:
        r = max(0.0, min(255.0, float(rgb_match.group(1))))
        g = max(0.0, min(255.0, float(rgb_match.group(2))))
        b = max(0.0, min(255.0, float(rgb_match.group(3))))
        alpha = float(rgb_match.group(4)) if rgb_match.group(4) is not None else 1.0
        if alpha <= 0.0:
            return None
        if alpha < 1.0:
            r = r * alpha + 255.0 * (1.0 - alpha)
            g = g * alpha + 255.0 * (1.0 - alpha)
            b = b * alpha + 255.0 * (1.0 - alpha)
        return f"{round(r):02X}{round(g):02X}{round(b):02X}"
    return None


def _parse_css_font_size(value: Optional[str]) -> Optional[float]:
    if not value:
        return None
    text = value.strip().lower()
    match = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)\s*(px|pt)", text)
    if not match:
        return None
    amount = float(match.group(1))
    unit = match.group(2)
    return round(amount * 0.75, 2) if unit == "px" else round(amount, 2)


def _parse_css_border_side(value: Optional[str]) -> Optional[Side]:
    if not value:
        return None
    text = value.strip().lower()
    if not text or "none" in text or re.search(r"\b0(?:px|pt)?\b", text):
        return None
    style = "thin"
    if "double" in text:
        style = "double"
    elif "dashed" in text:
        style = "dashed"
    elif "dotted" in text:
        style = "dotted"
    else:
        width_match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*px", text)
        if width_match:
            px = float(width_match.group(1))
            if px <= 0:
                return None
            if px >= 3:
                style = "thick"
            elif px >= 2:
                style = "medium"
    color_match = re.search(
        r"(rgba?\([^)]+\)|#[0-9a-f]{3,6}|\b(?:" + "|".join(_NAMED_COLORS) + r")\b)",
        text,
    )
    color_hex = _parse_css_color(color_match.group(1)) if color_match else None
    return Side(style=style, color=f"FF{color_hex}" if color_hex else None)


class _HTMLTableStyleParser(HTMLParser):
    """Extract per-cell inline CSS style maps aligned to the expanded HTML table grid."""

    def __init__(self, column_count: int) -> None:
        super().__init__()
        self.column_count = column_count
        self.rows: List[List[Dict[str, str]]] = []
        self._in_tbody = False
        self._has_tbody = False
        self._in_row = False
        self._tr_style: Dict[str, str] = {}
        self._current_row: List[Dict[str, str]] = []
        self._current_col = 0
        self._active_rowspans: Dict[int, Tuple[int, Dict[str, str]]] = {}

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        tag = tag.lower()
        attr_map = {k.lower(): (v or "") for k, v in attrs}
        if tag == "tbody":
            self._in_tbody = True
            self._has_tbody = True
            return
        if tag == "tr":
            if self._has_tbody and not self._in_tbody:
                return
            self._in_row = True
            self._tr_style = _parse_css_declarations(attr_map.get("style", ""))
            self._current_row = [{} for _ in range(self.column_count)]
            self._current_col = 0
            next_spans: Dict[int, Tuple[int, Dict[str, str]]] = {}
            for col, (remaining, inherited_style) in self._active_rowspans.items():
                if col < self.column_count:
                    self._current_row[col] = dict(inherited_style)
                if remaining > 1:
                    next_spans[col] = (remaining - 1, inherited_style)
            self._active_rowspans = next_spans
            return
        if tag in ("td", "th") and self._in_row:
            while (
                self._current_col < self.column_count
                and self._current_row[self._current_col]
            ):
                self._current_col += 1
            col_span = max(1, int(attr_map.get("colspan") or 1))
            row_span = max(1, int(attr_map.get("rowspan") or 1))
            cell_style = dict(self._tr_style)
            cell_style.update(_parse_css_declarations(attr_map.get("style", "")))
            for col in range(
                self._current_col, min(self._current_col + col_span, self.column_count)
            ):
                self._current_row[col] = dict(cell_style)
                if row_span > 1:
                    self._active_rowspans[col] = (row_span - 1, dict(cell_style))
            self._current_col += col_span

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "tbody":
            self._in_tbody = False
        elif tag == "tr" and self._in_row:
            self.rows.append(self._current_row)
            self._in_row = False


class ExcelTablePaster:
    """Write captured balance HTML tables directly into the workbook via openpyxl.

    Historically this class pasted HTML through the Windows system clipboard and
    Excel COM, which caused frequent clipboard and RPC failures. It now writes
    values, numeric types, number formats, row heights, and inline CSS styles
    directly with openpyxl (zero COM, zero clipboard).
    """

    COM_RETRY_ATTEMPTS = 40
    COM_RETRY_DELAY_SECONDS = 0.25
    FILE_ACCESS_RETRY_ATTEMPTS = 8
    FILE_ACCESS_RETRY_DELAY_SECONDS = 1
    RPC_E_CALL_REJECTED = -2147418111
    COM_DISCONNECTED_ERRORS = {-2147417848, -2147023174}

    AMOUNT_COLUMNS = {
        "BO_merchant_balance": (3, 6),
        "BO_provider_balance": (5, 8),
    }
    AMOUNT_NUMBER_FORMAT = "#,##0.00;\\-#,##0.00;0"
    HELPER_FORMAT_COLUMNS = {
        "BO_merchant_balance": (9, 21),
        "BO_provider_balance": (11, 20),
    }
    HELPER_NUMBER_FORMAT_CANONICAL = '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
    HELPER_NUMBER_FORMAT_LOCAL = '_(* # ##0_);_(* (# ##0);_(* "-"??_);_(@_)'
    FIRST_FORMULA_ROW = 6
    DEFAULT_ROW_HEIGHT = 13

    @staticmethod
    def _cell_matrix(
        capture: BalanceTableCapture, *, normalize_line_breaks: bool = False
    ) -> Tuple[Tuple[str, ...], ...]:
        matrix = []
        active_rowspans: Dict[int, int] = {}
        for row in capture.rows:
            cells = row.get("cells", [])
            spans = row.get("spans", [])
            expanded_cells = [""] * capture.column_count
            occupied_columns = set(active_rowspans)
            next_rowspans = {
                column: remaining - 1
                for column, remaining in active_rowspans.items()
                if remaining > 1
            }
            column = 0

            for index, cell in enumerate(cells):
                while column in occupied_columns:
                    column += 1
                span = spans[index] if index < len(spans) else {}
                column_span = max(1, span.get("col_span", 1))
                row_span = max(1, span.get("row_span", 1))
                if column < len(expanded_cells):
                    expanded_cells[column] = (
                        " ".join(cell.split())
                        if normalize_line_breaks
                        else cell
                    )

                for occupied_column in range(
                    column, min(column + column_span, len(expanded_cells))
                ):
                    occupied_columns.add(occupied_column)
                    if row_span > 1:
                        next_rowspans[occupied_column] = max(
                            next_rowspans.get(occupied_column, 0), row_span - 1
                        )
                column += column_span

            matrix.append(tuple(expanded_cells))
            active_rowspans = next_rowspans

        return tuple(matrix)

    @classmethod
    def _plain_text(cls, capture: BalanceTableCapture) -> str:
        return "\r\n".join(
            "\t".join(row)
            for row in cls._cell_matrix(capture, normalize_line_breaks=True)
        )

    @classmethod
    def _retry_excel_call(cls, operation):
        """Retained for backwards-compatibility with callers/tests referencing COM retry logic."""
        try:
            import pywintypes
            com_error_type = pywintypes.com_error
        except ImportError:
            com_error_type = Exception

        for attempt in range(cls.COM_RETRY_ATTEMPTS):
            try:
                return operation()
            except com_error_type as exc:
                if not exc.args or exc.args[0] != cls.RPC_E_CALL_REJECTED:
                    raise
                if attempt == cls.COM_RETRY_ATTEMPTS - 1:
                    raise
                time.sleep(cls.COM_RETRY_DELAY_SECONDS)

    @staticmethod
    def _parse_amount(value: str):
        text = value.strip().replace("\u00a0", "").replace("\u202f", "").replace(" ", "")
        if not text:
            return None
        if text in {"—", "–", "-"}:
            return text

        if "," in text and "." in text:
            text = text.replace(",", "")
        elif "," in text:
            if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", text):
                text = text.replace(",", "")
            else:
                text = text.replace(",", ".")

        try:
            return float(Decimal(text))
        except InvalidOperation as exc:
            raise ValueError(
                f"Unexpected balance amount from Backoffice: {value!r}."
            ) from exc

    @staticmethod
    def _parse_html_styles(capture: BalanceTableCapture) -> List[List[Dict[str, str]]]:
        if not capture.html or capture.column_count <= 0:
            return []
        parser = _HTMLTableStyleParser(capture.column_count)
        try:
            parser.feed(capture.html)
            parser.close()
        except Exception:
            return []
        return parser.rows

    @staticmethod
    def _apply_css_style(cell: Any, css: Dict[str, str], wrap_text: bool) -> None:
        if not css and not wrap_text:
            return

        # Font
        font_kwargs: Dict[str, Any] = {}
        weight = css.get("font-weight", "").strip().lower()
        if weight == "bold" or (weight.isdigit() and int(weight) >= 600):
            font_kwargs["bold"] = True
        if css.get("font-style", "").strip().lower() == "italic":
            font_kwargs["italic"] = True
        font_size = _parse_css_font_size(css.get("font-size"))
        if font_size is not None:
            font_kwargs["size"] = font_size
        font_family = css.get("font-family", "").split(",")[0].strip().strip("'\"")
        if font_family:
            font_kwargs["name"] = font_family
        font_color = _parse_css_color(css.get("color"))
        if font_color and font_color != "000000":
            font_kwargs["color"] = f"FF{font_color}"
        if font_kwargs:
            base = cell.font
            cell.font = Font(
                name=font_kwargs.get("name", base.name),
                size=font_kwargs.get("size", base.size),
                bold=font_kwargs.get("bold", base.bold),
                italic=font_kwargs.get("italic", base.italic),
                color=font_kwargs.get("color", base.color),
            )

        # Fill
        bg_hex = _parse_css_color(css.get("background-color"))
        if bg_hex:
            cell.fill = PatternFill(
                fill_type="solid",
                start_color=f"FF{bg_hex}",
                end_color=f"FF{bg_hex}",
            )

        # Alignment
        text_align = css.get("text-align", "").strip().lower()
        horizontal = {
            "left": "left",
            "center": "center",
            "right": "right",
            "justify": "justify",
        }.get(text_align)
        v_align = css.get("vertical-align", "").strip().lower()
        vertical = {
            "top": "top",
            "middle": "center",
            "bottom": "bottom",
        }.get(v_align)
        if horizontal or vertical or wrap_text:
            cell.alignment = Alignment(
                horizontal=horizontal,
                vertical=vertical,
                wrap_text=True if wrap_text else None,
            )

        # Borders
        border_sides = {}
        for side_name in ("top", "right", "bottom", "left"):
            side = _parse_css_border_side(css.get(f"border-{side_name}"))
            if side is not None:
                border_sides[side_name] = side
        if border_sides:
            cell.border = Border(**border_sides)

    @classmethod
    def _unmerge_intersecting(
        cls,
        worksheet: Any,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
    ) -> None:
        to_unmerge = [
            str(cell_range)
            for cell_range in list(worksheet.merged_cells.ranges)
            if (
                cell_range.min_row <= max_row
                and cell_range.max_row >= min_row
                and cell_range.min_col <= max_col
                and cell_range.max_col >= min_col
            )
        ]
        for range_ref in to_unmerge:
            worksheet.unmerge_cells(range_ref)

    @classmethod
    def paste_tables(
        cls,
        workbook_path: str,
        tables: Dict[str, Tuple[BalanceTableCapture, str]],
    ) -> List[str]:
        """Write all captured balance tables directly to the workbook using openpyxl."""
        if not tables:
            return []

        workbook = None
        for attempt in range(cls.FILE_ACCESS_RETRY_ATTEMPTS):
            try:
                workbook = openpyxl.load_workbook(
                    workbook_path, read_only=False, data_only=False
                )
                break
            except PermissionError as exc:
                if attempt == cls.FILE_ACCESS_RETRY_ATTEMPTS - 1:
                    raise PermissionError(
                        "Windows kept denying access to the ETL working workbook: "
                        f"{workbook_path}. Close it in Excel and wait for OneDrive "
                        "to finish syncing, then retry."
                    ) from exc
                time.sleep(cls.FILE_ACCESS_RETRY_DELAY_SECONDS)

        assert workbook is not None
        try:
            for sheet_name, (capture, start_cell) in tables.items():
                worksheet = workbook[sheet_name]
                start = worksheet[start_cell]
                top_row, left_column = start.row, start.column
                right_column = left_column + capture.column_count - 1
                pasted_last_row = top_row + capture.row_count - 1

                if sheet_name == "BO_provider_balance" and top_row == 6:
                    for col in range(1, 9):
                        worksheet.cell(5, col).value = None
                    for col in range(11, 21):
                        worksheet.cell(5, col).value = None

                merged_max_row = max(
                    (
                        cell_range.max_row
                        for cell_range in worksheet.merged_cells.ranges
                        if cell_range.min_col <= right_column
                        and cell_range.max_col >= left_column
                    ),
                    default=0,
                )
                data_max_row = 0
                for row_idx in range(top_row, worksheet.max_row + 1):
                    for col_idx in range(left_column, right_column + 1):
                        cell = worksheet.cell(row_idx, col_idx)
                        if cell.value is not None or (
                            cell.border
                            and any(
                                getattr(cell.border, side)
                                and getattr(cell.border, side).style is not None
                                for side in ("top", "bottom", "left", "right")
                            )
                        ):
                            data_max_row = row_idx
                            break

                last_existing_row = max(pasted_last_row, merged_max_row, data_max_row)

                # Unmerge any existing merged cells in the target + stale tail region
                cls._unmerge_intersecting(
                    worksheet, top_row, last_existing_row, left_column, right_column
                )

                # Clear target + stale tail region cleanly (values, borders, fills, alignments)
                empty_border = Border(
                    top=Side(style=None),
                    bottom=Side(style=None),
                    left=Side(style=None),
                    right=Side(style=None),
                )
                empty_fill = PatternFill(fill_type=None)
                default_alignment = Alignment()
                for row_idx in range(top_row, last_existing_row + 1):
                    for col_idx in range(left_column, right_column + 1):
                        cell = worksheet.cell(row_idx, col_idx)
                        cell.value = None
                        cell.border = copy(empty_border)
                        cell.fill = copy(empty_fill)
                        cell.alignment = copy(default_alignment)
                        cell.number_format = "General"

                matrix = cls._cell_matrix(capture)
                html_styles = cls._parse_html_styles(capture)
                amount_columns = cls.AMOUNT_COLUMNS.get(sheet_name)

                for row_offset, row_values in enumerate(matrix):
                    excel_row = top_row + row_offset
                    raw_row = capture.rows[row_offset] if row_offset < len(capture.rows) else {}
                    row_classes = raw_row.get("row_class", "").split()
                    is_header_row = "headers-row" in row_classes
                    style_row = (
                        html_styles[row_offset] if row_offset < len(html_styles) else []
                    )

                    for col_offset, raw_text in enumerate(row_values):
                        excel_col = left_column + col_offset
                        cell = worksheet.cell(excel_row, excel_col)
                        css = style_row[col_offset] if col_offset < len(style_row) else {}

                        is_amount_col = (
                            amount_columns is not None
                            and amount_columns[0] <= col_offset < amount_columns[1]
                        )

                        if is_amount_col and not is_header_row:
                            if str(raw_text).strip():
                                cell.value = cls._parse_amount(str(raw_text))
                            else:
                                cell.value = None
                            cell.number_format = cls.AMOUNT_NUMBER_FORMAT
                        else:
                            cell.value = raw_text if raw_text != "" else None

                        wrap_text = isinstance(raw_text, str) and (
                            "\n" in raw_text or "\r" in raw_text
                        )
                        cls._apply_css_style(cell, css, wrap_text)

                        # Keep amounts right-aligned unless CSS explicitly overrides
                        if is_amount_col and not is_header_row and not css.get("text-align"):
                            cell.alignment = Alignment(
                                horizontal="right",
                                vertical=cell.alignment.vertical,
                                wrap_text=cell.alignment.wrap_text,
                            )

                # Apply canonical helper number format to helper columns (rows 5..pasted_last_row)
                helper_columns = cls.HELPER_FORMAT_COLUMNS.get(sheet_name)
                if helper_columns and pasted_last_row >= cls.FIRST_FORMULA_ROW:
                    for row_idx in range(5, pasted_last_row + 1):
                        for col_idx in range(helper_columns[0], helper_columns[1] + 1):
                            worksheet.cell(
                                row_idx, col_idx
                            ).number_format = cls.HELPER_NUMBER_FORMAT_CANONICAL

                # Ensure amount/helper columns have enough width to display formatted numbers
                # on fresh worksheets without altering custom widths on existing workbooks.
                if amount_columns is not None:
                    for col_offset in range(amount_columns[0], amount_columns[1]):
                        col_letter = get_column_letter(left_column + col_offset)
                        if col_letter not in worksheet.column_dimensions:
                            worksheet.column_dimensions[col_letter].width = 22
                if helper_columns and pasted_last_row >= cls.FIRST_FORMULA_ROW:
                    for col_idx in range(helper_columns[0], helper_columns[1] + 1):
                        col_letter = get_column_letter(col_idx)
                        if col_letter not in worksheet.column_dimensions:
                            worksheet.column_dimensions[col_letter].width = 16

                # Normalize row height across the written and cleared tail rows
                if sheet_name in ("BO_merchant_balance", "BO_provider_balance"):
                    for row_idx in range(top_row, last_existing_row + 1):
                        worksheet.row_dimensions[row_idx].height = cls.DEFAULT_ROW_HEIGHT

            workbook.save(workbook_path)
            return []
        finally:
            workbook.close()
