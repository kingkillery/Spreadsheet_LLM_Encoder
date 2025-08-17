import openpyxl
from openpyxl.cell.cell import Cell
import re

# Regex for email validation
EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")


def infer_cell_data_type(cell: openpyxl.cell.cell.Cell) -> str:
    """
    Infers the data type of a cell based on its openpyxl data_type and value.
    """
    if cell.value is None:
        return "empty"

    # Check for email format using regex on the string value first
    if isinstance(cell.value, str) and EMAIL_REGEX.match(cell.value):
        return "email"

    data_type = cell.data_type
    if data_type == 's':
        return "text"
    elif data_type == 'n':
        return "numeric"
    elif data_type == 'b':
        return "boolean"
    elif data_type == 'd':
        return "datetime"
    elif data_type == 'e':
        return "error"
    elif data_type == 'f':
        if cell.value is not None:
            if isinstance(cell.value, str):
                return "text"
            elif isinstance(cell.value, (int, float)):
                return "numeric"
            elif isinstance(cell.value, bool):
                return "boolean"
            else:
                return "formula_cached_value"
        return "formula"
    elif data_type == 'g':
        if cell.value is not None:
            return "text"
        else:
            return "empty"
    else:
        return "unknown"


def categorize_number_format(number_format_string: str, cell: Cell) -> str:
    """
    Categorizes the number format of a cell, using the cell itself for context.
    """
    cell_data_type = infer_cell_data_type(cell)
    if cell_data_type not in ["numeric", "datetime"]:
        return "not_applicable"

    if number_format_string is None or number_format_string.lower() == "general":
        if cell_data_type == "datetime":
            return "datetime_general"
        return "general"

    if number_format_string == "@" or number_format_string.lower() == "text":
        return "text_format"

    if any(c in number_format_string for c in ['$', '€', '£', '¥']):
        return "currency"

    if '%' in number_format_string:
        return "percentage"

    if 'E+' in number_format_string or 'E-' in number_format_string.upper():
        return "scientific"

    if '#' in number_format_string and '/' in number_format_string and '?' in number_format_string:
        return "fraction"

    is_date_format = False
    is_time_format = False
    nf_lower = number_format_string.lower()
    date_keywords = ['yyyy', 'yy', 'mmmm', 'mmm', 'mm', 'dddd', 'ddd', 'dd', 'd']
    if any(keyword in nf_lower for keyword in date_keywords):
        is_date_format = True

    time_keywords = ['hh', 'h', 'ss', 's', 'am/pm', 'a/p']
    if any(keyword in nf_lower for keyword in time_keywords):
        is_time_format = True

    if ':' in number_format_string:
        temp_nf = number_format_string.replace('0', '').replace('#', '').replace(',', '').replace('.', '')
        if ':' in temp_nf:
            is_time_format = True

    if is_date_format and is_time_format:
        return "datetime_custom"
    elif is_date_format:
        return "date_custom"
    elif is_time_format:
        return "time_custom"

    if cell_data_type == "numeric":
        if number_format_string in ["0", "#,##0"]:
            return "integer"
        if number_format_string in ["0.00", "#,##0.00", "0.0", "#,##0.0"]:
            return "float"
        return "other_numeric"

    if cell_data_type == "datetime":
        return "other_date"

    return "unknown_format_category"


def get_number_format_string(cell: Cell) -> str:
    """Return the raw number format string for a cell."""
    try:
        nfs = cell.number_format
        if nfs is None or nfs == "":
            return "General"
        return str(nfs)
    except Exception:
        return "General"


def detect_semantic_type(cell: Cell) -> str:
    """Infer a higher level semantic type using number format and cell value."""
    data_type = infer_cell_data_type(cell)
    if data_type == "email":
        return "email"

    nfs = get_number_format_string(cell)
    category = categorize_number_format(nfs, cell)
    nfs_lower = nfs.lower()

    if category == "percentage":
        return "percentage"
    if category == "currency":
        return "currency"
    if category in ["date_custom", "datetime_custom", "datetime_general", "other_date"]:
        if ("yyyy" in nfs_lower or "yy" in nfs_lower) and not any(x in nfs_lower for x in ["m", "d"]):
            return "year"
        return "date"
    if category == "time_custom":
        return "time"
    if category == "scientific":
        return "scientific_notation"

    if data_type == "numeric":
        if isinstance(cell.value, int) or category == "integer":
            return "integer"
        if isinstance(cell.value, float) or category == "float":
            return "float"
        return "numeric"  # Fallback

    return data_type
