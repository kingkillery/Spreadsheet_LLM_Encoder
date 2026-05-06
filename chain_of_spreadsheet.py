"""
Chain of Spreadsheet (CoS) methodology for SpreadsheetLLM.
Implements the CoS pipeline structure as described in arXiv:2407.09025.

Configure an LLM backend via :func:`configure_backend` (preferred) or by
monkey-patching ``chain_of_spreadsheet._call_llm``. Example::

    import chain_of_spreadsheet as cos
    from llm_backend import OpenAIBackend

    cos.configure_backend(OpenAIBackend(model="gpt-4o-mini"))

Or, equivalently::

    cos._call_llm = lambda p: my_llm_client.complete(p)
"""
from __future__ import annotations

import json
import logging
import re
from typing import Dict, Iterable, List, Optional, Sequence

import paper_serializers
from llm_backend import LLMBackend  # noqa: F401  (re-exported for typing)
from tokenizer import count_tokens

logger = logging.getLogger(__name__)


# --- LLM backend hook --------------------------------------------------------

_BACKEND: Optional[LLMBackend] = None


def configure_backend(backend: Optional[LLMBackend]) -> None:
    """Register an :class:`LLMBackend` (or any ``Callable[[str], str]``).

    Pass ``None`` to clear the registered backend.
    """
    global _BACKEND
    _BACKEND = backend


def _call_llm(prompt: str) -> str:
    """Dispatch ``prompt`` to the configured backend.

    Returns ``_BACKEND(prompt)`` if a backend has been configured via
    :func:`configure_backend`; otherwise raises ``NotImplementedError``.

    Tests typically monkey-patch this function directly
    (e.g. ``chain_of_spreadsheet._call_llm = lambda p: ...``); when patched
    the backend mechanism is bypassed entirely.
    """
    if _BACKEND is not None:
        return _BACKEND(prompt)
    raise NotImplementedError(
        "_call_llm is a placeholder and has not been configured. "
        "Assign a real LLM callable before using the CoS pipeline:\n\n"
        "    import chain_of_spreadsheet as cos\n"
        "    cos._call_llm = lambda prompt: my_llm_client.complete(prompt)\n"
        "\nor:\n\n"
        "    from llm_backend import OpenAIBackend\n"
        "    cos.configure_backend(OpenAIBackend())\n"
    )


# --- Prompt Templates from Appendix L.3 -------------------------------------

QA_STAGE1_PROMPT_TEMPLATE = """
INSTRUCTION:
Given an input that is a string denoting data of cells in a table. The input table contains many tuples, describing the cells with content in the spreadsheet. Each tuple consists of two elements separated by a '|': the cell content and the cell address/region, like (Year|A1), ( |A1) or (IntNum|A1:B3). The content in some cells such as '#,##0'/'d-mmm-yy'/'H:mm:ss',etc., represents the CELL DATA FORMATS of Excel. The content in some cells such as 'IntNum'/'DateData'/'EmailData',etc., represents a category of data with the same format and similar semantics. For example, 'IntNum' represents integer type data, and 'ScientificNum' represents scientific notation type data. 'A1:B3' represents a region in a spreadsheet, from the first row to the third row and from column A to column B. Some cells with empty content in the spreadsheet are not entered. How many tables are there in the spreadsheet? Below is a question about one certain table in this spreadsheet. I need you to determine in which table the answer to the following question can be found, and return the RANGE of the ONE table you choose, LIKE ['range': 'A1:F9']. DON’T ADD OTHER WORDS OR EXPLANATION.

INPUT:
[Encoded Spreadsheet with compression]
[Question]
"""

QA_STAGE2_PROMPT_TEMPLATE = """
INSTRUCTION:
Given an input that is a string denoting data of cells in a table and a question about this table. The answer to the question can be found in the table. The input table includes many pairs, and each pair consists of a cell address and the text in that cell with a ',' in between, like 'A1,Year'. Cells are separated by '|' like 'A1,Year|A2,Profit'. The text can be empty so the cell data is like 'A1, |A2,Profit'. The cells are organized in row-major order. The answer to the input question is contained in the input table and can be represented by cell address. I need you to find the cell address of the answer in the given table based on the given question description, and return the cell ADDRESS of the answer like '[B3]' or '[SUM(A2:A10)]'. DON’T ADD ANY OTHER WORDS.

INPUT:
[Encoded Spreadsheet without compression]
[Question]
"""

QA_FINAL_SYNTHESIS_PROMPT_TEMPLATE = """
INSTRUCTION:
Given a spreadsheet question and candidate answers produced from row chunks of
the same table, choose or synthesize the final answer. Return only the final
answer in the same bracketed format, like '[B3]' or '[SUM(A2:A10)]'.
DON'T ADD ANY OTHER WORDS.

INPUT:
Question:
[Question]

Candidate Answers:
[Candidate Answers]
"""


# --- Internal helpers --------------------------------------------------------

# Match quoted or bare ranges — LLMs use all three forms freely.
_RANGE_RE = re.compile(r"""(?:['"]([A-Z]+\d+:[A-Z]+\d+)['"]|\b([A-Z]+\d+:[A-Z]+\d+)\b)""")
_STAGE2_LEGACY_WARNED = False


def _extract_ranges(text: str) -> List[str]:
    """Extract every Excel range from a Stage 1 response, preserving order."""
    ranges: List[str] = []
    seen = set()
    for match in _RANGE_RE.finditer(text or ""):
        rng = match.group(1) or match.group(2)
        if rng not in seen:
            ranges.append(rng)
            seen.add(rng)
    return ranges


def _build_stage2_prompt(prompt_input: str, query: str) -> str:
    prompt = QA_STAGE2_PROMPT_TEMPLATE.replace(
        "[Encoded Spreadsheet without compression]", prompt_input
    )
    return prompt.replace("[Question]", query)


# --- Stage 1 -----------------------------------------------------------------

def identify_table(
    encoding: Dict,
    query: str,
    sheet_name: Optional[str] = None,
) -> Optional[str]:
    """Identify the most relevant table for ``query`` (CoS Stage 1).

    Builds a paper-faithful compressed prompt from the selected sheet, calls
    the LLM, and parses the returned range.
    """
    if sheet_name is None:
        sheet_name = find_relevant_sheet(encoding, query)
    if not sheet_name:
        logger.warning("Could not identify a relevant sheet for the query.")
        return None

    sheets = encoding.get("sheets", {})
    if sheet_name not in sheets:
        logger.warning("Sheet '%s' not present in encoding.", sheet_name)
        return None
    sheet_data = sheets[sheet_name]

    prompt_input = paper_serializers.to_paper_compressed_prompt(
        sheet_data,
        coord_map=sheet_data.get("coord_map"),
    )

    prompt = QA_STAGE1_PROMPT_TEMPLATE.replace(
        "[Encoded Spreadsheet with compression]", prompt_input
    )
    prompt = prompt.replace("[Question]", query)

    llm_response = _call_llm(prompt)

    ranges = _extract_ranges(llm_response)
    if ranges:
        return ranges[0]

    logger.warning("Could not parse table range from LLM response: %s", llm_response)
    return None


def identify_tables(
    encoding: Dict,
    query: str,
    sheet_name: Optional[str] = None,
) -> List[str]:
    """Identify one or more relevant table ranges for ``query``.

    The paper prompt asks for one selected range, but model responses often
    contain structured lists. This helper keeps those ranges available for
    callers that want to inspect or post-process multiple predictions while
    :func:`identify_table` remains the single-range compatibility API.
    """
    if sheet_name is None:
        sheet_name = find_relevant_sheet(encoding, query)
    if not sheet_name:
        logger.warning("Could not identify a relevant sheet for the query.")
        return []

    sheets = encoding.get("sheets", {})
    if sheet_name not in sheets:
        logger.warning("Sheet '%s' not present in encoding.", sheet_name)
        return []
    sheet_data = sheets[sheet_name]

    prompt_input = paper_serializers.to_paper_compressed_prompt(
        sheet_data,
        coord_map=sheet_data.get("coord_map"),
    )

    prompt = QA_STAGE1_PROMPT_TEMPLATE.replace(
        "[Encoded Spreadsheet with compression]", prompt_input
    )
    prompt = prompt.replace("[Question]", query)
    return _extract_ranges(_call_llm(prompt))


# --- Sheet selection ---------------------------------------------------------

def find_relevant_sheet(encoding: Dict, query: str) -> Optional[str]:
    """Find the most relevant sheet for ``query``.

    Paper-faithful multi-sheet runs should use an LLM backend to select the
    relevant sheet. Keyword matching is retained as a documented fallback for
    no-backend runs or invalid LLM selections. Single-sheet fallback is always
    preserved.
    """
    sheet_names = list(encoding.get("sheets", {}))
    if len(sheet_names) == 1:
        return sheet_names[0]

    if _BACKEND is not None and len(sheet_names) > 1:
        try:
            picked = _llm_pick_sheet(sheet_names, query)
            if picked in sheet_names:
                return picked
            logger.warning(
                "LLM sheet selection returned %r; falling back to keyword matching.",
                picked,
            )
        except Exception as exc:  # pragma: no cover - defensive
            logger.warning("LLM sheet selection failed: %s; falling back to keywords.", exc)

    logger.info("Using keyword sheet selection fallback.")
    query_tokens = {t.lower() for t in query.split()}
    best_score = 0
    best_sheet: Optional[str] = None
    tied = False

    for name, sheet_data in encoding.get("sheets", {}).items():
        score = 0
        for value in sheet_data.get("cells", {}):
            lower_val = str(value).lower()
            if any(token in lower_val for token in query_tokens):
                score += 1
        if score > best_score:
            best_score = score
            best_sheet = name
            tied = False
        elif score == best_score and score > 0:
            tied = True

    if best_sheet and not tied:
        return best_sheet

    return best_sheet


def _llm_pick_sheet(sheet_names: Sequence[str], query: str) -> Optional[str]:
    listing = ", ".join(sheet_names)
    prompt = (
        "You are picking the most relevant spreadsheet tab for a question.\n"
        f"Available sheets: {listing}\n"
        f"Question: {query}\n"
        "Respond with ONLY the exact sheet name and no other text."
    )
    answer = _call_llm(prompt).strip()
    # Strip optional surrounding quotes
    answer = answer.strip("'\"")
    return answer if answer in sheet_names else None


def _find_relevant_sheet(encoding: Dict, query: str) -> Optional[str]:
    """Deprecated wrapper; use :func:`find_relevant_sheet` directly."""
    return find_relevant_sheet(encoding, query)


# --- Stage 2 -----------------------------------------------------------------

def generate_response(
    sheet_data: Dict,
    query: str,
    *,
    workbook_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    table_range: Optional[str] = None,
    coord_map: Optional[Dict] = None,
) -> str:
    """Generate a Stage 2 response for ``query`` over the identified table.

    When ``workbook_path``, ``sheet_name`` and ``table_range`` are all
    supplied, the paper-faithful uncompressed sub-range is read directly
    from the original workbook (Section 4.2). Compact ranges produced by
    the encoder's coordinate remapping are unmapped first.

    Otherwise the function falls back to passing the (compressed)
    ``sheet_data`` into the prompt and emits a one-time warning that this
    is not paper-faithful.
    """
    payload = build_stage2_prompt_payload(
        sheet_data,
        query,
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        table_range=table_range,
        coord_map=coord_map,
    )
    return _call_llm(payload["prompt"])


def build_stage2_prompt_payload(
    sheet_data: Dict,
    query: str,
    *,
    workbook_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    table_range: Optional[str] = None,
    coord_map: Optional[Dict] = None,
) -> Dict[str, Optional[str]]:
    """Build a Stage 2 prompt plus mode metadata.

    Returns a dict containing ``prompt``, ``prompt_input``, ``stage2_mode``,
    ``table_range`` and ``original_range``. ``stage2_mode`` is
    ``"original_workbook_uncompressed"`` for the paper-faithful path and
    ``"compressed_json_fallback"`` for the legacy fallback.
    """
    global _STAGE2_LEGACY_WARNED

    if workbook_path and sheet_name and table_range:
        if coord_map is None:
            coord_map = sheet_data.get("coord_map") if isinstance(sheet_data, dict) else None

        original_range = table_range
        if coord_map:
            unmapped = paper_serializers.unremap_range(table_range, coord_map)
            if unmapped is not None:
                original_range = unmapped
            else:
                logger.warning(
                    "unremap_range returned None for %s; using compact range as-is.",
                    table_range,
                )

        prompt_input = paper_serializers.to_stage2_uncompressed_prompt(
            workbook_path, sheet_name, original_range
        )
        stage2_mode = "original_workbook_uncompressed"
    else:
        if not _STAGE2_LEGACY_WARNED:
            logger.warning(
                "generate_response called without workbook_path/sheet_name/"
                "table_range; using compressed encoding (paper specifies an "
                "uncompressed sub-range for Stage 2)."
            )
            _STAGE2_LEGACY_WARNED = True
        prompt_input = json.dumps(sheet_data, ensure_ascii=False)
        original_range = None
        stage2_mode = "compressed_json_fallback"

    prompt = _build_stage2_prompt(prompt_input, query)
    return {
        "prompt": prompt,
        "prompt_input": prompt_input,
        "stage2_mode": stage2_mode,
        "table_range": table_range,
        "original_range": original_range,
    }


# --- Algorithm 2: row-chunked QA --------------------------------------------

def table_split_qa(
    sheet_data: Dict,
    table_range: str,
    query: str,
    *,
    workbook_path: Optional[str] = None,
    sheet_name: Optional[str] = None,
    coord_map: Optional[Dict] = None,
    token_limit: int = 4096,
    header_rows: Optional[Iterable[int]] = None,
    tokenizer_model: str = "gpt-4",
) -> str:
    """Handle QA for large tables via Algorithm 2 (Appendix M.2) row chunking.

    Real path (when ``workbook_path`` and ``sheet_name`` are provided):

    1. Resolve the compact ``table_range`` to original-workbook coordinates
       via ``coord_map`` (taken from ``sheet_data`` if not passed).
    2. Default ``header_rows`` to ``[r1]`` if not specified.
    3. Compute the full uncompressed prompt's token count; if it fits
       ``token_limit``, delegate to :func:`generate_response` once.
    4. Otherwise greedily partition body rows so each
       ``header + chunk`` rendering fits ``token_limit`` (using
       :func:`tokenizer.count_tokens`), guaranteeing forward progress with
       at least one body row per chunk.
    5. Substitute ``header + chunk`` pairs into ``QA_STAGE2_PROMPT_TEMPLATE``,
       call :func:`_call_llm` per chunk, and aggregate.

    Legacy path (no workbook info): emits a deprecation warning and calls
    :func:`generate_response` once on the supplied ``sheet_data``.
    """
    if not (workbook_path and sheet_name):
        logger.warning(
            "table_split_qa called without workbook_path/sheet_name; "
            "row-chunking requires the original workbook. Falling back to a "
            "single generate_response call. This legacy path is deprecated."
        )
        return generate_response(sheet_data, query)

    if coord_map is None and isinstance(sheet_data, dict):
        coord_map = sheet_data.get("coord_map")

    original_range = table_range
    if coord_map:
        unmapped = paper_serializers.unremap_range(table_range, coord_map)
        if unmapped is not None:
            original_range = unmapped
        else:
            logger.warning(
                "unremap_range returned None for %s; using compact range as-is.",
                table_range,
            )

    r1, c1, r2, c2 = paper_serializers.parse_range(original_range)

    header_list: List[int] = (
        [r1] if header_rows is None else sorted({int(r) for r in header_rows})
    )
    body_rows: List[int] = [r for r in range(r1, r2 + 1) if r not in header_list]

    full_prompt_input = paper_serializers.to_stage2_uncompressed_prompt(
        workbook_path, sheet_name, original_range
    )
    full_prompt = _build_stage2_prompt(full_prompt_input, query)

    if count_tokens(full_prompt, model=tokenizer_model) <= token_limit:
        return generate_response(
            sheet_data,
            query,
            workbook_path=workbook_path,
            sheet_name=sheet_name,
            table_range=table_range,
            coord_map=coord_map,
        )

    logger.info("Table is too large; applying Algorithm 2 row chunking.")

    answers: List[str] = []
    i = 0
    n = len(body_rows)
    while i < n:
        # Always include at least one body row; greedily extend while the
        # rendered prompt fits within token_limit.
        chunk: List[int] = [body_rows[i]]
        candidate_prompt = _render_chunk_prompt(
            workbook_path, sheet_name, original_range,
            header_list + chunk, query,
        )
        candidate_tokens = count_tokens(candidate_prompt, model=tokenizer_model)
        if candidate_tokens > token_limit:
            # Single header+row already exceeds the budget. We can't split
            # further without losing the header context, so emit a warning
            # and dispatch as-is rather than hang or skip silently.
            logger.warning(
                "Header + row %s alone is %d tokens (> token_limit=%d); "
                "sending anyway. Consider raising token_limit or shrinking "
                "the column band.",
                body_rows[i], candidate_tokens, token_limit,
            )
        j = i + 1
        while j < n:
            tentative = chunk + [body_rows[j]]
            tentative_prompt = _render_chunk_prompt(
                workbook_path, sheet_name, original_range,
                header_list + tentative, query,
            )
            if count_tokens(tentative_prompt, model=tokenizer_model) > token_limit:
                break
            chunk = tentative
            candidate_prompt = tentative_prompt
            j += 1
        answers.append(_call_llm(candidate_prompt))
        i = j

    return _synthesize_chunk_answers(query, answers)


def _render_chunk_prompt(
    workbook_path: str,
    sheet_name: str,
    original_range: str,
    rows: Sequence[int],
    query: str,
) -> str:
    pair_string = paper_serializers.stage2_pairs_from_rows(
        workbook_path, sheet_name, original_range, rows
    )
    return _build_stage2_prompt(pair_string, query)


def _synthesize_chunk_answers(query: str, answers: Sequence[str]) -> str:
    """Run the final CoS synthesis step over per-chunk candidate answers.

    Skips the synthesis LLM call when there is at most one candidate: zero
    answers degenerate to ``"[]"`` and a single answer is already final, so
    funneling it back through the LLM only risks a reformat or hallucinated
    rewrite while paying for an extra round trip.
    """
    if not answers:
        return "[]"
    if len(answers) == 1:
        return answers[0]
    answer_block = "\n".join(f"{idx + 1}. {answer}" for idx, answer in enumerate(answers))
    prompt = QA_FINAL_SYNTHESIS_PROMPT_TEMPLATE.replace("[Question]", query)
    prompt = prompt.replace("[Candidate Answers]", answer_block)
    return _call_llm(prompt)
