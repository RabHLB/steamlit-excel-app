import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from datetime import datetime
import os

# Define the Excel file path
file_path = "example.xlsx"

# Ensure the Excel file exists
if not os.path.exists(file_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Column A", "Column B", "Timestamp"])  # Add headers
    wb.save(file_path)

# Load the data
def load_data():
    return pd.read_excel(file_path, engine="openpyxl")

# Function to add data
def add_data(col_a_input, col_b_input):
    if not col_a_input or not col_b_input:
        st.warning("Please enter values for both fields.")
        return

    wb = load_workbook(file_path)
    ws = wb.active
    next_row = ws.max_row + 1
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ws[f"A{next_row}"] = col_a_input
    ws[f"B{next_row}"] = col_b_input
    ws[f"C{next_row}"] = timestamp

    wb.save(file_path)
    wb.close()
    st.success("Data added successfully!")

# Streamlit UI
st.title("Excel Data Entry & Viewer")

# Input fields
col_a = st.text_input("Enter value for Column A:")
col_b = st.text_input("Enter value for Column B:")

if st.button("Add to Excel"):
    add_data(col_a, col_b)

# Display data
st.subheader("Current Data in Excel:")
df = load_data()
st.dataframe(df)
