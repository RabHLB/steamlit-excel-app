import streamlit as st
from openpyxl import load_workbook, Workbook
import pandas as pd
from datetime import datetime
import os
from typing import List, Dict, Optional
import json


class ExcelApp:
    def __init__(self, file_path: str, headers: List[str]):
        """
        Initialize the Excel App
        Args:
            file_path (str): Path to the Excel file
            headers (List[str]): List of column headers
        """
        self.file_path = file_path
        self.headers = headers
        self.setup_page()
        self.initialize_excel()
        self.load_data()

    def setup_page(self) -> None:
        """Configure Streamlit page settings"""
        st.set_page_config(
            layout="wide",
            page_title="Excel Data Manager",
            page_icon="📊"
        )
        st.title("Excel Data Entry & Manager")

    def initialize_excel(self) -> None:
        """Initialize Excel file with required sheets"""
        try:
            if not os.path.exists(self.file_path):
                self.create_new_workbook()
            else:
                self.ensure_sheets_exist()
        except Exception as e:
            st.error(f"Error initializing Excel file: {str(e)}")

    def create_new_workbook(self) -> None:
        """Create a new Excel workbook with required sheets"""
        wb = Workbook()
        ws_main = wb.active
        ws_main.title = "Main"
        ws_main.append(self.headers)
        ws_log = wb.create_sheet("Log")
        ws_log.append(["Timestamp", "Action", "User", "Details"])
        wb.save(self.file_path)

    def ensure_sheets_exist(self) -> None:
        """Ensure required sheets exist in workbook"""
        wb = load_workbook(self.file_path)
        if "Main" not in wb.sheetnames:
            ws_main = wb.create_sheet("Main")
            ws_main.append(self.headers)
        if "Log" not in wb.sheetnames:
            ws_log = wb.create_sheet("Log")
            ws_log.append(["Timestamp", "Action", "User", "Details"])
        wb.save(self.file_path)

    def load_data(self) -> None:
        """Load data from Excel file into session state"""
        try:
            df = pd.read_excel(self.file_path, sheet_name="Main", engine="openpyxl")
            if "df" not in st.session_state:
                st.session_state.df = df.copy()
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")
            st.session_state.df = pd.DataFrame(columns=self.headers)

    def create_grid(self) -> None:
        """Create and configure the editable data grid"""
        st.subheader("📝 Data Editor")

        # Use Streamlit's native data editor
        edited_df = st.data_editor(
            st.session_state.df,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                col: st.column_config.NumberColumn(col, help=f"Enter {col} value")
                for col in self.headers[2:]  # Configure number columns for months
            },
            hide_index=True,
        )

        if edited_df is not None:
            st.session_state.df = edited_df.copy()

    def add_new_row(self) -> None:
        """Add a new empty row to the dataframe"""
        col1, col2 = st.columns([1, 5])
        with col1:
            if st.button("➕ Add Row"):
                new_row = pd.DataFrame([{col: None for col in st.session_state.df.columns}])
                st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                st.rerun()

    def save_data(self) -> None:
        """Save data to Excel file"""
        col1, col2 = st.columns([1, 5])
        with col1:
            if st.button("💾 Save"):
                if self.validate_data_for_save():
                    self.perform_save()
                    st.success("✅ Data saved successfully!")

    def validate_data_for_save(self) -> bool:
        """Validate data before saving"""
        if st.session_state.df.empty or st.session_state.df.iloc[:, 1:].isnull().all().all():
            st.error("❌ Cannot save: The sheet has no valid data!")
            return False
        return True

    def perform_save(self) -> None:
        """Perform the save operation"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = {
            "timestamp": timestamp,
            "action": "Update",
            "user": "System",
            "details": json.dumps({"rows_updated": len(st.session_state.df)})
        }

        with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as writer:
            st.session_state.df.to_excel(writer, sheet_name="Main", index=False)

            # Update log
            wb = writer.book
            ws_log = wb["Log"]
            ws_log.append([log_entry["timestamp"], log_entry["action"],
                           log_entry["user"], log_entry["details"]])

    def clear_data(self) -> None:
        """Clear all data from the main sheet"""
        col1, col2 = st.columns([1, 5])
        with col1:
            if st.button("🗑️ Clear Data"):
                if st.session_state.df.empty:
                    st.warning("Sheet is already empty!")
                    return

                st.session_state.show_confirm = True

        if getattr(st.session_state, 'show_confirm', False):
            st.warning("⚠️ Are you sure you want to clear all data?")
            col1, col2 = st.columns([1, 1])
            with col1:
                if st.button("Yes, clear data"):
                    self.perform_clear()
                    st.session_state.show_confirm = False
                    st.rerun()
            with col2:
                if st.button("Cancel"):
                    st.session_state.show_confirm = False
                    st.rerun()

    def perform_clear(self) -> None:
        """Perform the clear operation"""
        st.session_state.df = pd.DataFrame(columns=self.headers)
        wb = load_workbook(self.file_path)
        ws_main = wb["Main"]

        # Keep header row and clear the rest
        for row in ws_main.iter_rows(min_row=2, max_row=ws_main.max_row):
            for cell in row:
                cell.value = None

        # Log the clear action
        ws_log = wb["Log"]
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws_log.append([timestamp, "Clear", "System", "Cleared all data"])
        wb.save(self.file_path)
        st.success("🧹 Data cleared successfully!")

    def show_log(self) -> None:
        """Display the change log"""
        st.subheader("📋 Change Log")
        try:
            df_log = pd.read_excel(self.file_path, sheet_name="Log", engine="openpyxl")
            st.dataframe(df_log, height=200)
        except Exception as e:
            st.error(f"Error loading log: {str(e)}")


def main():
    # Initialize session state
    if 'trigger_rerun' not in st.session_state:
        st.session_state.trigger_rerun = False

    # Define headers
    headers = [
        "Account Number", "Account Name",
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]

    # Create an instance of ExcelApp with proper initialization
    file_path = "example.xlsx"
    app = ExcelApp(file_path=file_path, headers=headers)

    # Create sidebar with options
    with st.sidebar:
        st.header("Options")
        st.markdown("---")
        if st.button("📖 Show Instructions"):
            st.info("""
            **Instructions:**
            1. Edit cells directly by clicking
            2. Use Tab or Enter to move between cells
            3. Add new rows using the Add Row button
            4. Save changes using the Save button
            5. View change history in the Log section
            """)

    # Main content area
    app.create_grid()

    # Action buttons
    st.markdown("---")
    app.add_new_row()
    app.save_data()
    app.clear_data()

    # Show log at the bottom
    st.markdown("---")
    app.show_log()


if __name__ == "__main__":
    main()