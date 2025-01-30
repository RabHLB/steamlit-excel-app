import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from openpyxl import load_workbook, Workbook
import pandas as pd
from datetime import datetime

# Set Streamlit to use full-width mode
st.set_page_config(layout="wide")

# Title of the app
st.title("Excel Data Entry & Viewer with Full-Width Editable Table")

# File path for the Excel file
file_path = "example.xlsx"

# Ensure the Excel file has headers
headers = [
    "Time Stamp", "Account Number", "Account Name",
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

try:
    wb = load_workbook(file_path)
    ws = wb.active
    if ws.max_row == 0:  # If the file is empty
        ws.append(headers)  # Add headers
        wb.save(file_path)
except FileNotFoundError:
    # Create a new workbook if the file doesn't exist
    wb = Workbook()
    ws = wb.active
    ws.append(headers)  # Add headers
    wb.save(file_path)

# Read data from the Excel file
try:
    df = pd.read_excel(file_path, engine="openpyxl")
except FileNotFoundError:
    df = pd.DataFrame(columns=headers)  # Create an empty DataFrame with headers

# Editable grid/table with full-width auto-resizing
st.subheader("Editable Data Table:")

gb = GridOptionsBuilder.from_dataframe(df)

# Enable full-width auto-sizing and resizable columns
gb.configure_default_column(editable=True, resizable=True)
gb.configure_grid_options(domLayout='autoHeight', autoSizeColumns=True)
gb.configure_columns(df.columns, flex=1)  # Make all columns stretch evenly

# Display the grid with full width
grid_response = AgGrid(
    df,
    gridOptions=gb.build(),
    update_mode=GridUpdateMode.VALUE_CHANGED,
    fit_columns_on_grid_load=True,
    allow_unsafe_jscode=True,
    theme="streamlit",  # Use Streamlit theme
    height=800,  # Allows more rows to be visible
)

# Updated DataFrame after editing
updated_df = grid_response["data"]

# Button to save the data
if st.button("Save to Excel"):
    # Add current timestamp for any new rows
    updated_df["Time Stamp"] = updated_df["Time Stamp"].fillna(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    # Save the DataFrame to the Excel file
    updated_df.to_excel(file_path, index=False, engine="openpyxl")
    st.success("Data saved successfully!")

# Display the updated data
st.subheader("Updated Data in Excel:")
st.dataframe(updated_df)
