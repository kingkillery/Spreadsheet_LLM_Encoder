import os
import pandas as pd
import tempfile
import json

from app_helpers import (
    analyze_sheet_for_compression_insights,
    format_key_number_format,
    get_format_regions,
)

# Helper function as defined in the streamlit_app.py logic (or a simplified version for testing)


def read_csv_with_multiple_encodings(file_path):
    """
    Tries to read a CSV file using multiple encodings (utf-8, latin-1, iso-8859-1).
    Raises the last exception if all attempts fail.
    Returns a pandas DataFrame if successful.
    """
    try:
        df = pd.read_csv(file_path, encoding='utf-8')
        return df
    except UnicodeDecodeError:
        # print("UTF-8 failed, trying latin-1") # Optional: for debugging test execution
        try:
            df = pd.read_csv(file_path, encoding='latin-1')
            return df
        except UnicodeDecodeError:
            # print("latin-1 failed, trying iso-8859-1") # Optional: for debugging test execution
            try:
                df = pd.read_csv(file_path, encoding='iso-8859-1')
                return df
            except Exception as e:
                raise e  # Raise the last exception (iso-8859-1 read error)
        except Exception as e:
            raise e  # Raise the last exception (latin-1 read error)
    except Exception as e:
        raise e  # Raise the last exception (utf-8 read error)


# Sample CSV data encoded in 'latin-1'
sample_latin1_csv_data = "Name,City\nJules,Paris\nRené,Montréal\nBjörn,Göteborg".encode('latin-1')


def test_csv_encoding_handling():
    """
    Tests the CSV encoding handling by attempting to read a latin-1 encoded CSV.
    """
    # Create a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".csv", mode='wb') as tmp_file:
        tmp_file.write(sample_latin1_csv_data)
        tmp_file_path = tmp_file.name

    try:
        df = read_csv_with_multiple_encodings(tmp_file_path)

        assert df is not None, "DataFrame should not be None"
        assert not df.empty, "DataFrame should not be empty"

        # Check for expected data (adjust based on actual data and column names)
        # Example: Check if "René" is in the 'Name' column
        assert "René" in df["Name"].values, "Expected name 'René' not found in DataFrame"
        assert "Björn" in df["Name"].values, "Expected name 'Björn' not found in DataFrame"
        assert "Montréal" in df["City"].values, "Expected city 'Montréal' not found in DataFrame"
        assert "Göteborg" in df["City"].values, "Expected city 'Göteborg' not found in DataFrame"

        # Check shape
        assert df.shape == (3, 2), f"DataFrame shape mismatch. Expected (3, 2), got {df.shape}"

    finally:
        # Clean up the temporary file
        if os.path.exists(tmp_file_path):
            os.remove(tmp_file_path)


def test_format_helpers_accept_current_and_legacy_keys():
    current_key = json.dumps({"type": "integer", "nfs": "#,##0"}, sort_keys=True)
    legacy_key = json.dumps(
        {"font": {"bold": True}, "number_format": "0.00"},
        sort_keys=True,
    )

    current = {"formats": {current_key: ["A1:A2"]}}
    legacy = {"format_regions": {legacy_key: ["B1:B2"]}}

    assert get_format_regions(current) == current["formats"]
    assert get_format_regions(legacy) == legacy["format_regions"]
    assert format_key_number_format(json.loads(current_key)) == "#,##0"
    assert format_key_number_format(json.loads(legacy_key)) == "0.00"

    insights = analyze_sheet_for_compression_insights(current)
    assert insights["format_analysis"]["num_unique_formats_overall"] == 1

# Example of how to run this test with pytest (if desired, not run by the agent directly)
# if __name__ == "__main__":
#     pytest.main([__file__])
