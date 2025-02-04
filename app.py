import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from openpyxl import load_workbook, Workbook
import pandas as pd
from datetime import datetime

# Set Streamlit to use full-width mode
st.set_page_config(layout="wide")

# Title of the app
st.title("Excel Data Entry & Viewer with Change Log & Clear Function")

# File path for the Excel file
file_path = "example.xlsx"

# Define headers for the main sheet
headers = [
    "Account Number", "Account Name", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Ensure the Excel file has the required sheets
try:
    wb = load_workbook(file_path)
    if "Main" not in wb.sheetnames:
        ws_main = wb.create_sheet("Main")
        ws_main.append(headers)  # Add headers
    if "Log" not in wb.sheetnames:
        ws_log = wb.create_sheet("Log")
        ws_log.append(["Timestamp", "Action", "Updated Data"])
    wb.save(file_path)
except FileNotFoundError:
    # Create a new workbook if the file doesn't exist
    wb = Workbook()
    ws_main = wb.active
    ws_main.title = "Main"
    ws_main.append(headers)  # Add headers
    ws_log = wb.create_sheet("Log")
    ws_log.append(["Timestamp", "Action", "Updated Data"])
    wb.save(file_path)

# Read data from the "Main" sheet
try:
    df = pd.read_excel(file_path, sheet_name="Main", engine="openpyxl")
except FileNotFoundError:
    df = pd.DataFrame(columns=headers)  # Create an empty DataFrame with headers

# Initialize the table in session state
if "df" not in st.session_state:
    st.session_state.df = df.copy()  # Store the DataFrame in session state

# Editable grid/table with full-width auto-resizing
st.subheader("Editable Data Table:")
gb = GridOptionsBuilder.from_dataframe(st.session_state.df)
gb.configure_default_column(editable=True, resizable=True)
gb.configure_grid_options(
    domLayout='autoHeight',
    autoSizeColumns=True,
    enterMovesDownAfterEdit=True,  # Allows TAB or ENTER to move to the next cell
    suppressRowTransform=True,  # Prevents unwanted resets
    singleClickEdit=True,  # Start editing on a single click
)
gb.configure_columns(st.session_state.df.columns, flex=1)

# Render the editable table
grid_response = AgGrid(
    st.session_state.df,
    gridOptions=gb.build(),
    update_mode=GridUpdateMode.VALUE_CHANGED,  # Update only when values change
    fit_columns_on_grid_load=True,
    allow_unsafe_jscode=True,
    theme="streamlit",
    height=800,
)

# Process updated data from the grid
if grid_response and "data" in grid_response:
    updated_df = pd.DataFrame(grid_response["data"])  # Capture the updated data

    # Validate numeric columns
    numeric_columns = headers[2:]  # Columns 3 to 14
    invalid_entries = False
    for col in numeric_columns:
        updated_df[col] = updated_df[col].apply(
            lambda x: x if pd.isna(x) or (str(x).isdigit() or isinstance(x, (int, float))) else "INVALID"
        )
        if "INVALID" in updated_df[col].values:
            invalid_entries = True

    if invalid_entries:
        st.error("Only numbers are allowed in the columns: January to December. Please correct the invalid entries.")
    else:
        st.session_state.df = updated_df.copy()  # Save changes to session state

# Add a button to insert a blank row
if st.button("Add New Row"):
    new_row = pd.DataFrame([{col: None for col in st.session_state.df.columns}])
    st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
    # Trigger rerun without using deprecated APIs
    st.session_state["trigger"] = not st.session_state.get("trigger", False)


if st.button("Save to Excel"):
    # Validation: Check if the sheet contains valid data
    if st.session_state.df.empty or st.session_state.df.iloc[:, 1:].isnull().all().all():
        st.error("Cannot save: The sheet has no valid data! Please enter at least one value.")
    else:
        # Ensure no empty rows or cells replace valid data
        for col in st.session_state.df.columns[1:]:
            st.session_state.df[col] = st.session_state.df[col].apply(
                lambda x: x if pd.notna(x) and str(x).strip() != "" else None
            )

        # Save the data from session state
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = [timestamp, "Update", st.session_state.df.to_json()]

        # Load workbook and save updates
        wb = load_workbook(file_path)
        ws_main = wb["Main"]
        ws_log = wb["Log"]

        # Save the updated main sheet
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            st.session_state.df.to_excel(writer, sheet_name="Main", index=False)

        # Append log entry to log sheet
        ws_log.append(log_entry)
        wb.save(file_path)

        st.success("Data saved successfully! Changes logged.")

# Display the Log Sheet
st.subheader("Change Log:")
try:
    df_log = pd.read_excel(file_path, sheet_name="Log", engine="openpyxl")
    st.dataframe(df_log)
except FileNotFoundError:
    st.error("Log sheet not found.")

# Button to clear all data from the Main sheet (leaving headers intact)
if st.button("Clear Main Sheet"):
    wb = load_workbook(file_path)
    ws_main = wb["Main"]

    # Delete all rows except the header row
    for row in ws_main.iter_rows(min_row=2, max_row=ws_main.max_row):
        for cell in row:
            cell.value = None

    wb.save(file_path)
    st.success("Main sheet cleared! Only headers remain.")
