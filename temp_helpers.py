import openpyxl # Required for type hinting Cell

# --- Helper Functions for Format Information Extraction (Step 1) ---
def infer_cell_data_type(cell: openpyxl.cell.cell.Cell) -> str:
    """
    Infers the data type of a cell based on its openpyxl data_type and value.
    """
    if cell.value is None:
        return "empty"

    data_type = cell.data_type
    if data_type == 's': # Standard string
        return "text"
    elif data_type == 'n': # Number
        return "numeric"
    elif data_type == 'b': # Boolean
        return "boolean"
    elif data_type == 'd': # Datetime
        return "datetime"
    elif data_type == 'e': # Error
        return "error"
    elif data_type == 'f': # Formula
        if cell.value is not None: # openpyxl might store the last calculated value
            if isinstance(cell.value, str):
                return "text"
            elif isinstance(cell.value, (int, float)):
                return "numeric"
            elif isinstance(cell.value, bool):
                return "boolean"
            # Dates from formulas could be datetime objects (type 'd') or numbers formatted as dates (type 'n').
            # This function primarily uses cell.data_type, so 'd' is handled. 'n' will be "numeric".
            else:
                return "formula_cached_value" # Other cached types
        return "formula" # No cached value or type not easily inferred
    elif data_type == 'g': # General (often for empty or untyped cells)
        # Already checked for cell.value is None. If 'g' and not None, treat as text.
        if cell.value is not None:
            return "text"
        else:
            return "empty" # Fallback, though None check should catch this.
    else: # Includes 'is' (inlineStr), 'str' (formula string, rare for type)
        return "unknown"

def categorize_number_format(number_format_string: str, cell_data_type: str) -> str:
    """
    Categorizes the number format of a cell.
    """
    if cell_data_type not in ["numeric", "datetime"]:
        return "not_applicable"

    if number_format_string is None or number_format_string.lower() == "general":
        if cell_data_type == "datetime":
            # If openpyxl identified it as datetime but format is general, it's a date.
            return "datetime_general"
        return "general" # General numeric

    # Text format override (Format Code: @)
    if number_format_string == "@" or number_format_string.lower() == "text":
        return "text_format"

    # Currency symbols (common ones)
    if any(c in number_format_string for c in ['$', '€', '£', '¥']):
        return "currency"

    # Percentage symbol
    if '%' in number_format_string:
        return "percentage"

    # Scientific notation (e.g., 1.23E+05)
    if 'E+' in number_format_string or 'E-' in number_format_string.upper(): # Openpyxl stores 'E+' or 'e+'
        return "scientific"

    # Fraction (e.g., # ?/?)
    if '#' in number_format_string and '/' in number_format_string and '?' in number_format_string:
        return "fraction"

    # Date/Time related checks
    # These patterns are indicative and might need refinement for complex/custom Excel formats.
    is_date_format = False
    is_time_format = False

    # Common date components (lowercase for matching)
    nf_lower = number_format_string.lower()
    date_keywords = ['yyyy', 'yy', 'mmmm', 'mmm', 'mm', 'dddd', 'ddd', 'dd', 'd']
    if any(keyword in nf_lower for keyword in date_keywords):
        is_date_format = True

    # Common time components (lowercase for matching)
    time_keywords = ['hh', 'h', 'ss', 's', 'am/pm', 'a/p']
    if any(keyword in nf_lower for keyword in time_keywords):
        is_time_format = True

    # Presence of colons, not part of a numeric placeholder like 0.00 or #,##0.00
    if ':' in number_format_string:
        # Avoid misclassifying "0.00" or "#,##0.00" if they contain ':' somehow (unlikely in valid formats)
        # A simple check: if there are digits around colons, it's more likely time.
        # More robustly, check if it's not purely a number format with colons.
        temp_nf = number_format_string.replace('0','').replace('#','').replace(',','').replace('.','')
        if ':' in temp_nf: # If colon persists after removing numeric placeholders
            is_time_format = True


    if is_date_format and is_time_format:
        return "datetime_custom" # Contains both date and time elements
    elif is_date_format:
        if 'yyyy' in nf_lower or 'mmmm' in nf_lower or 'dddd' in nf_lower:
            return "date_long" # Longer date formats
        elif 'yy' in nf_lower or 'mmm' in nf_lower or 'ddd' in nf_lower:
            return "date_short" # Shorter date formats
        return "other_date" # Other formats that look like dates
    elif is_time_format:
        return "time_custom" # Formats that look like times

    # Fallback for numeric types if no other category matched yet
    if cell_data_type == "numeric":
        # Check for common explicit numeric formats that aren't "General"
        if number_format_string in ["0", "#,##0", "0.00", "#,##0.00", "0.0", "#,##0.0"]:
            return "general_numeric_explicit"
        return "other_numeric" # Numeric, but not currency, percentage, scientific, fraction, or common general.

    # If cell_data_type was 'datetime' but format string didn't trigger specific date/time patterns
    if cell_data_type == "datetime":
        return "other_date" # It's a date type, but format is unusual (e.g. a custom number format applied to a date)

    return "unknown_format_category" # Default if no category fits
