import streamlit as st
import pandas as pd
import json
import base64
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
import openpyxl # Added for CSV to Excel conversion
import tempfile # Will be used for CSV temp file
import os # Will be used for CSV temp file
from openpyxl.utils import get_column_letter # For table detection

def get_download_link(json_data, filename="spreadsheet_data.json"):
    """
    Generates an HTML download link for JSON data.

    Args:
        json_data (dict or list): The JSON data to be downloaded.
        filename (str, optional): The desired filename for the downloaded file.
                                  Defaults to "spreadsheet_data.json".

    Returns:
        str: An HTML string representing the download link.
    """
    json_str = json.dumps(json_data, indent=2)
    b64 = base64.b64encode(json_str.encode()).decode()
    href = f'<a href="data:file/json;base64,{b64}" download="{filename}">Download JSON File</a>'
    return href

# Simplified table detection function (as per prompt)
def detect_tables_in_sheet(sheet):
    """
    Detects a single table in a given Openpyxl worksheet object based on contiguous data.

    Args:
        sheet (openpyxl.worksheet.worksheet.Worksheet): The worksheet to analyze.

    Heuristic:
        Identifies the bounding box of all non-empty cells in the sheet.
        This bounding box is considered a single "table". This is a very basic
        heuristic and does not identify multiple tables or complex layouts.

    Returns:
        list: A list containing a single dictionary if a data block is found,
              representing the detected table. The dictionary has keys "range"
              (e.g., "A1:E10") and "method" (e.g., "basic_contiguous_block").
              Returns an empty list if the sheet is empty or no data is found.
    """
    if not sheet:
        return []

    min_r, min_c, max_r, max_c = sheet.min_row, sheet.min_column, sheet.max_row, sheet.max_column

    if sheet.max_row == 1 and sheet.max_column == 1 and sheet.cell(1,1).value is None: # Empty sheet
        return []

    # Find actual data bounds
    data_min_row, data_min_col = sheet.max_row + 1, sheet.max_column + 1
    data_max_row, data_max_col = sheet.min_row -1, sheet.min_column -1

    has_data = False
    for r_idx in range(min_r, max_r + 1): # Corrected variable name r to r_idx
        for c_idx in range(min_c, max_c + 1): # Corrected variable name c to c_idx
            if sheet.cell(row=r_idx, column=c_idx).value is not None:
                has_data = True
                data_min_row = min(data_min_row, r_idx)
                data_min_col = min(data_min_col, c_idx)
                data_max_row = max(data_max_row, r_idx)
                data_max_col = max(data_max_col, c_idx)

    if not has_data:
        return []

    table_range = f"{get_column_letter(data_min_col)}{data_min_row}:{get_column_letter(data_max_col)}{data_max_row}"
    return [{"range": table_range, "method": "basic_contiguous_block"}]

def main():
    """
    Main function to run the Streamlit application.
    Handles file uploads, processing (Excel/CSV), encoding via Spreadsheet_LLM_Encoder,
    basic table detection, and displays the resulting JSON.
    """
    st.title("Spreadsheet to Encoded JSON")
    st.write("Upload your spreadsheet file (Excel or CSV) and get JSON in return.")

    # K value slider
    k_value = st.slider("Neighborhood distance (k)", min_value=0, max_value=10, value=2, step=1)

    # File uploader
    uploaded_file = st.file_uploader("Choose a spreadsheet file", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        try:
            file_extension = uploaded_file.name.split(".")[-1].lower() # Ensure lowercase extension
            json_data = None
            # This variable will store the path to the Excel file (either original or CSV-converted)
            # that is used for encoding and table detection.
            processed_file_path_for_table_detection = None

            if file_extension in ['xlsx', 'xls']:
                # Handling for direct Excel file uploads
                # Create a temporary file to store the uploaded Excel data.
                # This path is then used by spreadsheet_llm_encode and for table detection.
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_extension}") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    processed_file_path_for_table_detection = tmp.name

                try:
                    # Encode the Excel file to JSON
                    json_data = spreadsheet_llm_encode(processed_file_path_for_table_detection, k=k_value)

                    # Perform table detection if encoding was successful and data exists
                    if json_data and "sheets" in json_data and processed_file_path_for_table_detection:
                        # Load the workbook (the one used by encoder) to get sheet objects for table detection
                        source_workbook = openpyxl.load_workbook(processed_file_path_for_table_detection, data_only=True)
                        for sheet_name_wb in source_workbook.sheetnames:
                            if sheet_name_wb in json_data["sheets"]: # Ensure sheet exists in encoded data
                                current_sheet_obj = source_workbook[sheet_name_wb]
                                # Detect tables in the current sheet
                                detected_tables_list = detect_tables_in_sheet(current_sheet_obj)
                                if detected_tables_list: # Add detected tables to the JSON output
                                    json_data["sheets"][sheet_name_wb]["detected_tables"] = detected_tables_list
                finally:
                    # Clean up the temporary Excel file
                    if processed_file_path_for_table_detection and os.path.exists(processed_file_path_for_table_detection):
                        os.remove(processed_file_path_for_table_detection)

            elif file_extension == 'csv':
                # Handling for CSV file uploads
                st.write("Processing CSV: Converting to temporary Excel for encoding...")
                # Read CSV into pandas DataFrame
                df = pd.read_csv(uploaded_file)

                # Create an in-memory Excel workbook and populate it from the DataFrame
                wb = openpyxl.Workbook()
                sheet = wb.active
                sheet.title = "Sheet1" # Default sheet name for CSV converted data

                # Write headers from DataFrame
                for col_idx, column_name in enumerate(df.columns, 1):
                    sheet.cell(row=1, column=col_idx, value=str(column_name))
                # Write data rows from DataFrame
                for row_idx, row_data_tuple in enumerate(df.itertuples(index=False), 2):
                    for col_idx, cell_value in enumerate(row_data_tuple, 1):
                        sheet.cell(row=row_idx, column=col_idx, value=cell_value)

                # Create a temporary directory to store the converted Excel file
                temp_csv_conversion_dir = tempfile.mkdtemp()
                # Define the path for the temporary Excel file
                processed_file_path_for_table_detection = os.path.join(temp_csv_conversion_dir, "temp_for_csv.xlsx")

                try:
                    # Save the DataFrame content as a temporary Excel file
                    wb.save(processed_file_path_for_table_detection)
                    # Encode the temporary Excel file to JSON
                    json_data = spreadsheet_llm_encode(processed_file_path_for_table_detection, k=k_value)

                    # Perform table detection if encoding was successful and data exists
                    if json_data and "sheets" in json_data and processed_file_path_for_table_detection:
                        # Load the temporary workbook (the one used by encoder) for table detection
                        source_workbook = openpyxl.load_workbook(processed_file_path_for_table_detection, data_only=True)
                        for sheet_name_wb in source_workbook.sheetnames: # Should typically be "Sheet1"
                             if sheet_name_wb in json_data["sheets"]: # Ensure sheet exists in encoded data
                                current_sheet_obj = source_workbook[sheet_name_wb]
                                # Detect tables in the current sheet
                                detected_tables_list = detect_tables_in_sheet(current_sheet_obj)
                                if detected_tables_list: # Add detected tables to the JSON output
                                    json_data["sheets"][sheet_name_wb]["detected_tables"] = detected_tables_list
                finally:
                    # Clean up the temporary Excel file and directory
                    if processed_file_path_for_table_detection and os.path.exists(processed_file_path_for_table_detection):
                        os.remove(processed_file_path_for_table_detection)
                    if temp_csv_conversion_dir and os.path.exists(temp_csv_conversion_dir): # Ensure directory exists before trying to remove
                        os.rmdir(temp_csv_conversion_dir)

            # Display results or error message
            if json_data:
                # Prepare JSON string for display and download
                json_str = json.dumps(json_data, indent=2)

                # Display the JSON
                st.subheader("JSON Output")
                st.json(json_data)

                # Text area for easy copying
                st.subheader("Copy JSON")
                st.text_area("Copy this JSON:", value=json_str, height=250)

                # Download button
                st.markdown(get_download_link(json_data), unsafe_allow_html=True)
            else:
                st.error("Failed to process the spreadsheet and generate JSON data.")
            
        except Exception as e:
            st.error(f"Error processing file: {e}")

if __name__ == "__main__":
    main()
