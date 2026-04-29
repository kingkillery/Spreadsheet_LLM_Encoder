"""Example for the two-stage chain-of-spreadsheet pipeline."""
import sys
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
from chain_of_spreadsheet import find_relevant_sheet, identify_table, generate_response


def main():
    if len(sys.argv) < 3:
        print("Usage: python example_chain_usage.py <excel_path> <query>")
        return
    excel_path = sys.argv[1]
    query = " ".join(sys.argv[2:])

    encoding = spreadsheet_llm_encode(excel_path)
    if not encoding:
        print("Failed to encode spreadsheet")
        return

    sheet_name = find_relevant_sheet(encoding, query)
    table_range = identify_table(encoding, query)
    if not sheet_name or not table_range:
        print("Could not identify a relevant table")
        return

    sheet_data = encoding["sheets"][sheet_name]
    answer = generate_response(sheet_data, query)
    print(f"Selected sheet: {sheet_name}")
    print(f"Selected range: {table_range}")
    print(answer)


if __name__ == "__main__":
    main()
