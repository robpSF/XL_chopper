import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def create_excel_sheets(df, column):
    unique_values = df[column].unique()
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for value in unique_values:
            subset_df = df[df[column] == value]
            subset_df.to_excel(writer, sheet_name=str(value), index=False)
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

