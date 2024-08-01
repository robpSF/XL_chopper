import streamlit as st
import pandas as pd
import re
from io import BytesIO

def sanitize_sheet_name(name):
    sanitized = re.sub(r'[^0-9a-zA-Z_ ]+', '', str(name))[:31]
    return sanitized

def create_excel_sheets(df, column):
    unique_values = df[column].dropna().unique()  # Drop NaN values
    if len(unique_values) == 0:
        st.warning("No valid sheet names found. Check if the selected column has values.")
        return None
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        seen_sheet_names = set()
        for value in unique_values:
            sheet_name = sanitize_sheet_name(value)
            if sheet_name in seen_sheet_names:
                counter = 1
                new_sheet_name = f"{sheet_name}_{counter}"
                while new_sheet_name in seen_sheet_names:
                    counter += 1
                    new_sheet_name = f"{sheet_name}_{counter}"
                sheet_name = new_sheet_name
            seen_sheet_names.add(sheet_name)
            subset_df = df[df[column] == value]
            subset_df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

def combine_sheets_to_one(uploaded_file):
    try:
        all_dfs = pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        return None
    combined_df = pd.DataFrame()
    header = None
    for sheet_name, df in all_dfs.items():
        if header is None:
            header = df.columns
        else:
            if not header.equals(df.columns):
                st.warning(f"Headers do not match for sheet: {sheet_name}. Ignoring this sheet.")
                continue
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='Combined', index=False)
    output.seek(0)
    return output

st.title('Excel Sheet Splitter and Combiner')

mode = st.selectbox('Select Mode', ['Cut', 'Paste'])
if mode == 'Cut':
    st.write('Upload an Excel file, specify the column, and download a new Excel file with separate sheets for each unique value in the specified column.')
else:
    st.write('Upload an Excel file with multiple sheets and combine them into one sheet. If headers do not match, those sheets will be ignored.')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_file:
    if mode == 'Cut':
        df = pd.read_excel(uploaded_file)
        st.write('File successfully uploaded.')
        columns = df.columns.tolist()
        selected_column = st.selectbox('Select the column to split sheets by', columns)
        if selected_column:
            st.write(f'You selected: {selected_column}')
            if st.button('Create Excel'):
                output = create_excel_sheets(df, selected_column)
                if output:
                    st.download_button(label="Download Excel file", data=output, file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        if st.button('Create Excel'):
            output = combine_sheets_to_one(uploaded_file)
            if output:
                st.download_button(label="Download Combined Excel file", data=output, file_name="combined_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
