import pandas as pd
import tempfile
import os
import pytest # Assuming pytest will be used to run tests

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
                raise e # Raise the last exception (iso-8859-1 read error)
        except Exception as e:
            raise e # Raise the last exception (latin-1 read error)
    except Exception as e:
        raise e # Raise the last exception (utf-8 read error)

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

# Example of how to run this test with pytest (if desired, not run by the agent directly)
# if __name__ == "__main__":
#     pytest.main([__file__])
