import streamlit as st
import pandas as pd
import re
from io import BytesIO

def sanitize_sheet_name(name):
    # Remove invalid characters and truncate to 31 characters
    sanitized = re.sub(r'[^0-9a-zA-Z_ ]+', '', str(name))[:31]
    return sanitized

def create_excel_sheets(df, column):
    unique_values = df[column].unique()
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        seen_sheet_names = set()
        for value in unique_values:
            sheet_name = sanitize_sheet_name(value)
            if sheet_name in seen_sheet_names:
                # Ensure sheet names are unique
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

st.title('Excel Sheet Splitter')
st.write('Upload an Excel file, specify the column, and download a new Excel file with separate sheets for each unique value in the specified column.')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write('File successfully uploaded.')
    columns = df.columns.tolist()
    selected_column = st.selectbox('Select the column to split sheets by', columns)
    if selected_column:
        st.write(f'You selected: {selected_column}')
        if st.button('Create Excel'):
            output = create_excel_sheets(df, selected_column)
            st.download_button(label="Download Excel file", data=output, file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
