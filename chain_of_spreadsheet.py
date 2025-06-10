"""Simple two stage pipeline for SpreadsheetLLM encoded files."""
import json
from typing import Dict, Optional

def identify_table(encoding: Dict, query: str) -> Optional[str]:
    """Return the sheet name most relevant to the query.

    Parameters
    ----------
    encoding: dict
        SpreadsheetLLM JSON encoding produced by ``spreadsheet_llm_encode``.
    query: str
        Natural language question about the spreadsheet.

    Returns
    -------
    str or None
        Name of the sheet that best matches the query, or ``None`` if no match.
    """
    query_tokens = [t.lower() for t in query.split()]
    best_score = 0
    best_sheet = None

    for sheet_name, sheet_data in encoding.get("sheets", {}).items():
        score = 0
        for value in sheet_data.get("cells", {}):
            lower_val = str(value).lower()
            for token in query_tokens:
                if token in lower_val:
                    score += 1
        if score > best_score:
            best_score = score
            best_sheet = sheet_name
    return best_sheet

def generate_response(sheet_data: Dict, query: str) -> str:
    """Generate a text answer using the selected sheet.

    This minimal implementation simply lists cells that contain words
    from the query. It can be replaced by a call to an external LLM
    if desired.
    """
    matches = {}
    query_tokens = [t.lower() for t in query.split()]
    for value, cells in sheet_data.get("cells", {}).items():
        lower_val = str(value).lower()
        if any(token in lower_val for token in query_tokens):
            matches[value] = cells
    if not matches:
        return f"No relevant information found for '{query}'."
    lines = [f"Matches for '{query}':"]
    for val, refs in matches.items():
        lines.append(f"  Value '{val}' at {', '.join(refs)}")
    return "\n".join(lines)

