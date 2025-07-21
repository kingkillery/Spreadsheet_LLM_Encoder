import streamlit as st
import json
import base64
import logging
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode
import openpyxl
import tempfile
from app_helpers import (
    detect_tables_in_sheet,
    extract_chart_info_from_sheet,
    parse_number_format_string,
    extract_sheet_metadata,
    analyze_sheet_for_compression_insights,
    generate_common_value_map,
    parse_cell_ref,
    parse_range_ref,
    is_cell_within_parsed_range,
    read_csv_with_multiple_encodings,
)
import os

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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


def main():
    """
    Main function to run the Streamlit application.
    Handles file uploads, processing (Excel/CSV), encoding via Spreadsheet_LLM_Encoder,
    and various post-processing steps (table detection, chart extraction, number format parsing,
    sheet metadata extraction, compression insights, common value mapping, chart-table linking),
    and displays the resulting JSON.
    """
    st.title("Spreadsheet to Encoded JSON")
    st.write("Upload your spreadsheet file (Excel or CSV) and get JSON in return.")

    k_value = st.slider("Neighborhood distance (k)", min_value=0, max_value=10, value=2, step=1)
    uploaded_file = st.file_uploader("Choose a spreadsheet file", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        try:
            file_extension = uploaded_file.name.split(".")[-1].lower()
            json_data = None
            processed_file_path_for_postprocessing = None

            if file_extension in ['xlsx', 'xls']:
                with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_extension}") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    processed_file_path_for_postprocessing = tmp.name

                try:
                    with st.spinner("Encoding spreadsheet..."):
                        json_data = spreadsheet_llm_encode(processed_file_path_for_postprocessing, k=k_value)

                    if json_data and "sheets" in json_data and processed_file_path_for_postprocessing:
                        source_workbook = openpyxl.load_workbook(processed_file_path_for_postprocessing, data_only=True)
                        for sheet_name_wb in source_workbook.sheetnames:
                            if sheet_name_wb in json_data["sheets"]:
                                current_sheet_obj = source_workbook[sheet_name_wb]
                                sheet_data_node = json_data["sheets"][sheet_name_wb]

                                detected_tables_list = detect_tables_in_sheet(current_sheet_obj)
                                if detected_tables_list:
                                    sheet_data_node["detected_tables"] = detected_tables_list

                                chart_list = extract_chart_info_from_sheet(current_sheet_obj)
                                if chart_list:
                                    sheet_data_node["charts"] = chart_list

                                if "format_regions" in sheet_data_node:
                                    parsed_number_formats_map = {}
                                    for fmt_key_json, _ in sheet_data_node["format_regions"].items():
                                        try:
                                            fmt_details = json.loads(fmt_key_json)
                                            number_format_str = fmt_details.get("number_format")
                                            if number_format_str and number_format_str not in parsed_number_formats_map:
                                                parsed_number_formats_map[number_format_str] = parse_number_format_string(number_format_str)
                                        except json.JSONDecodeError:
                                            st.warning(f"Could not parse format key string: {fmt_key_json}")
                                    if parsed_number_formats_map:
                                        sheet_data_node["parsed_number_formats"] = parsed_number_formats_map

                                sheet_metadata = extract_sheet_metadata(current_sheet_obj)
                                sheet_data_node["sheet_level_metadata"] = sheet_metadata

                                compression_analysis_results = analyze_sheet_for_compression_insights(sheet_data_node)
                                sheet_data_node["compression_insights"] = compression_analysis_results

                                common_value_map_results = generate_common_value_map(sheet_data_node)
                                if common_value_map_results:
                                     sheet_data_node["common_value_map"] = common_value_map_results

                                # Chart to Table Linking
                                if "charts" in sheet_data_node and "detected_tables" in sheet_data_node:
                                    for chart_dict in sheet_data_node["charts"]:
                                        if chart_dict.get("anchor"):
                                            chart_anchor_col, chart_anchor_row = parse_cell_ref(chart_dict["anchor"])
                                            if chart_anchor_col and chart_anchor_row:
                                                linked_tables = []
                                                for table_dict in sheet_data_node["detected_tables"]:
                                                    if table_dict.get("full_range"):
                                                        r_start_col, r_start_row, r_end_col, r_end_row = parse_range_ref(table_dict["full_range"])
                                                        if r_start_col and is_cell_within_parsed_range(chart_anchor_col, chart_anchor_row, r_start_col, r_start_row, r_end_col, r_end_row):
                                                            linked_tables.append(table_dict["full_range"])
                                                if linked_tables:
                                                    chart_dict["linked_table_ranges"] = linked_tables
                finally:
                    if processed_file_path_for_postprocessing and os.path.exists(processed_file_path_for_postprocessing):
                        os.remove(processed_file_path_for_postprocessing)

            elif file_extension == 'csv':
                st.write("Processing CSV: Converting to temporary Excel for encoding...")
                df = read_csv_with_multiple_encodings(uploaded_file)
                json_data = None
                if df is None:
                    st.error("Error reading CSV file with supported encodings.")


                if df is not None:
                    wb = openpyxl.Workbook()
                    sheet = wb.active
                    sheet.title = "Sheet1"
                    for col_idx, column_name in enumerate(df.columns, 1):
                        sheet.cell(row=1, column=col_idx, value=str(column_name))
                    for row_idx, row_data_tuple in enumerate(df.itertuples(index=False), 2):
                        for col_idx, cell_value in enumerate(row_data_tuple, 1):
                            sheet.cell(row=row_idx, column=col_idx, value=cell_value)

                    temp_csv_conversion_dir = tempfile.mkdtemp()
                    processed_file_path_for_postprocessing = os.path.join(temp_csv_conversion_dir, "temp_for_csv.xlsx")

                    try:
                        wb.save(processed_file_path_for_postprocessing)
                        with st.spinner("Encoding spreadsheet..."):
                            json_data = spreadsheet_llm_encode(processed_file_path_for_postprocessing, k=k_value)

                        if json_data and "sheets" in json_data and processed_file_path_for_postprocessing:
                            source_workbook = openpyxl.load_workbook(processed_file_path_for_postprocessing, data_only=True)
                            for sheet_name_wb in source_workbook.sheetnames:
                                 if sheet_name_wb in json_data["sheets"]:
                                    current_sheet_obj = source_workbook[sheet_name_wb]
                                    sheet_data_node = json_data["sheets"][sheet_name_wb]

                                    detected_tables_list = detect_tables_in_sheet(current_sheet_obj)
                                    if detected_tables_list:
                                        sheet_data_node["detected_tables"] = detected_tables_list

                                    chart_list = extract_chart_info_from_sheet(current_sheet_obj)
                                    if chart_list:
                                        sheet_data_node["charts"] = chart_list

                                    if "format_regions" in sheet_data_node:
                                        parsed_number_formats_map = {}
                                        for fmt_key_json, _ in sheet_data_node["format_regions"].items():
                                            try:
                                                fmt_details = json.loads(fmt_key_json)
                                                number_format_str = fmt_details.get("number_format")
                                                if number_format_str and number_format_str not in parsed_number_formats_map:
                                                    parsed_number_formats_map[number_format_str] = parse_number_format_string(number_format_str)
                                            except json.JSONDecodeError:
                                                st.warning(f"Could not parse format key string: {fmt_key_json}")
                                        if parsed_number_formats_map:
                                            sheet_data_node["parsed_number_formats"] = parsed_number_formats_map

                                    sheet_metadata = extract_sheet_metadata(current_sheet_obj)
                                    sheet_data_node["sheet_level_metadata"] = sheet_metadata

                                    compression_analysis_results = analyze_sheet_for_compression_insights(sheet_data_node)
                                    sheet_data_node["compression_insights"] = compression_analysis_results

                                    common_value_map_results = generate_common_value_map(sheet_data_node)
                                    if common_value_map_results:
                                         sheet_data_node["common_value_map"] = common_value_map_results

                                    # Chart to Table Linking (for CSV, charts won't exist from original, but tables might be detected)
                                    if "charts" in sheet_data_node and "detected_tables" in sheet_data_node:
                                        for chart_dict in sheet_data_node["charts"]: # This list will be empty for CSVs
                                            if chart_dict.get("anchor"):
                                                chart_anchor_col, chart_anchor_row = parse_cell_ref(chart_dict["anchor"])
                                                if chart_anchor_col and chart_anchor_row:
                                                    linked_tables = []
                                                    for table_dict in sheet_data_node["detected_tables"]:
                                                        if table_dict.get("full_range"):
                                                            r_start_col, r_start_row, r_end_col, r_end_row = parse_range_ref(table_dict["full_range"])
                                                            if r_start_col and is_cell_within_parsed_range(chart_anchor_col, chart_anchor_row, r_start_col, r_start_row, r_end_col, r_end_row):
                                                                linked_tables.append(table_dict["full_range"])
                                                    if linked_tables:
                                                        chart_dict["linked_table_ranges"] = linked_tables
                    finally:
                        if processed_file_path_for_postprocessing and os.path.exists(processed_file_path_for_postprocessing):
                            os.remove(processed_file_path_for_postprocessing)
                        if 'temp_csv_conversion_dir' in locals() and temp_csv_conversion_dir and os.path.exists(temp_csv_conversion_dir):
                            os.rmdir(temp_csv_conversion_dir)
                # If df is None (all decoding failed), json_data remains None, and this block is skipped.
                # The existing error message for failed processing will be shown later if json_data is still None.

            if json_data:
                json_str = json.dumps(json_data, indent=2)
                st.subheader("JSON Output")
                st.json(json_data)
                st.subheader("Copy JSON")
                st.text_area("Copy this JSON:", value=json_str, height=250)
                st.markdown(get_download_link(json_data), unsafe_allow_html=True)
            else:
                st.error("Failed to process the spreadsheet and generate JSON data.")
            
        except Exception as e:
            st.error(f"Error processing file: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    main()
