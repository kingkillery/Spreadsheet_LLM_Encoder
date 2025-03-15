import streamlit as st
import pandas as pd
import json
import base64

def get_download_link(json_data, filename="spreadsheet_data.json"):
    """Generate a download link for the JSON data"""
    json_str = json.dumps(json_data, indent=2)
    b64 = base64.b64encode(json_str.encode()).decode()
    href = f'<a href="data:file/json;base64,{b64}" download="{filename}">Download JSON File</a>'
    return href

def main():
    st.title("Spreadsheet to Encoded JSON")
    st.write("Upload your spreadsheet file (Excel or CSV) and get JSON in return.")

    # File uploader
    uploaded_file = st.file_uploader("Choose a spreadsheet file", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        try:
            # Read the uploaded file
            file_extension = uploaded_file.name.split(".")[-1]
            
            if file_extension in ['xlsx', 'xls']:
                df = pd.read_excel(uploaded_file)
            elif file_extension == 'csv':
                df = pd.read_csv(uploaded_file)
            
            # Convert to JSON
            json_data = json.loads(df.to_json(orient="records"))
            json_str = json.dumps(json_data, indent=2)
            
            # Display the JSON
            st.subheader("JSON Output")
            st.json(json_data)
            
            # Text area for easy copying
            st.subheader("Copy JSON")
            st.text_area("Copy this JSON:", value=json_str, height=250)
            
            # Download button
            st.markdown(get_download_link(json_data), unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"Error processing file: {e}")

if __name__ == "__main__":
    main()
