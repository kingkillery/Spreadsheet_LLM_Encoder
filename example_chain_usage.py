"""End-to-end Chain-of-Spreadsheet pipeline demonstration.

This example encodes a spreadsheet, identifies a relevant table, and generates
a response using the CoS pipeline. It uses EchoBackend by default (returns a
canned response) for deterministic testing without external LLM calls. If the
OPENAI_API_KEY environment variable is set, uses OpenAIBackend instead.

Usage:
    python example_chain_usage.py <excel_path> <query>

Example:
    python example_chain_usage.py data.xlsx "What was the profit in 2021?"
"""
import os
import sys
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
from chain_of_spreadsheet import (
    find_relevant_sheet,
    identify_table,
    generate_response,
    configure_backend,
)
from llm_backend import EchoBackend, OpenAIBackend
import paper_serializers


def main():
    if len(sys.argv) < 3:
        print("Usage: python example_chain_usage.py <excel_path> <query>")
        sys.exit(1)

    excel_path = sys.argv[1]
    query = " ".join(sys.argv[2:])

    print(f"Encoding spreadsheet: {excel_path}")
    encoding = spreadsheet_llm_encode(excel_path)
    if not encoding:
        print("Failed to encode spreadsheet")
        sys.exit(1)

    # Configure LLM backend: use OpenAI if API key is set, otherwise EchoBackend
    if os.getenv("OPENAI_API_KEY"):
        print("Using OpenAIBackend (gpt-4o-mini)")
        configure_backend(OpenAIBackend(model="gpt-4o-mini"))
    else:
        print(
            "OPENAI_API_KEY not set; using EchoBackend with canned responses "
            "(for demo only)"
        )
        configure_backend(
            EchoBackend(response="['range': 'A1:B10']")
        )

    # Stage 1: Find relevant sheet
    sheet_name = find_relevant_sheet(encoding, query)
    if not sheet_name:
        print("Could not identify a relevant sheet")
        sys.exit(1)

    print(f"Selected sheet: {sheet_name}")

    # Stage 1: Identify table range
    table_range = identify_table(encoding, query, sheet_name=sheet_name)
    if not table_range:
        print("Could not identify a relevant table")
        sys.exit(1)

    print(f"Selected range (compact): {table_range}")

    # Get sheet data and coord_map for Stage 2
    sheet_data = encoding["sheets"][sheet_name]
    coord_map = sheet_data.get("coord_map")

    # Unmap compact range back to original workbook coordinates if needed
    original_range = table_range
    if coord_map:
        unmapped = paper_serializers.unremap_range(table_range, coord_map)
        if unmapped:
            original_range = unmapped
            print(f"Selected range (original): {original_range}")

    # Stage 2: Generate response using paper-faithful uncompressed prompt
    # Pass workbook_path, sheet_name, table_range to trigger the real
    # (paper-aligned) Stage 2 path that reads from the original workbook.
    answer = generate_response(
        sheet_data,
        query,
        workbook_path=excel_path,
        sheet_name=sheet_name,
        table_range=table_range,
        coord_map=coord_map,
    )

    print(f"\nAnswer:")
    print(answer)


if __name__ == "__main__":
    main()
