
import tempfile
import os
import openpyxl
from Spreadsheet_LLM_Encoder import spreadsheet_llm_encode


def create_sample_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = 'Header'
    ws['A2'] = 42
    ws['B2'] = 'Data'
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


def test_spreadsheet_llm_encode_structure():
    excel_path = create_sample_excel()
    try:
        encoding = spreadsheet_llm_encode(excel_path, k=1)
        assert 'sheets' in encoding
        assert encoding['sheets'], 'No sheets found'
        sheet = next(iter(encoding['sheets'].values()))
        assert 'cells' in sheet
        assert 'formats' in sheet
    finally:
        os.remove(excel_path)
