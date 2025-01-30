import streamlit as st
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime

# Title of the app
st.title("Excel Data Entry & Viewer")

# Input fields for data entry
col_a = st.text_input("Enter value for Column A:")
col_b = st.text_input("Enter value for Column B:")

# Button to add data to the Excel file
if st.button("Add to Excel"):
    file_path = "example.xlsx"  # Path to the Excel file

    # Load the workbook and select the active sheet
    try:
        wb = load_workbook(file_path)
        ws = wb.active
    except FileNotFoundError:
        st.error("The Excel file does not exist. Please create 'example.xlsx' in the directory.")

    # Find the next empty row
    next_row = ws.max_row + 1

    # Add the user input and timestamp to the Excel file
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws[f"A{next_row}"] = col_a
    ws[f"B{next_row}"] = col_b
    ws[f"C{next_row}"] = timestamp

    # Save the updated workbook
    wb.save(file_path)
    st.success("Data added successfully!")

# Display the contents of the Excel file
st.subheader("Current Data in Excel:")

try:
    df = pd.read_excel("example.xlsx", engine="openpyxl")
    st.dataframe(df)
except FileNotFoundError:
    st.error("The Excel file does not exist. Please create 'example.xlsx' in the directory.")
