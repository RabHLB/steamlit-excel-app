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
        ws_log.append(["Timestamp", "Action", "Updated Data"])
        wb.save(self.file_path)

    def ensure_sheets_exist(self) -> None:
        """Ensure required sheets exist in workbook"""
        wb = load_workbook(self.file_path)
        if "Main" not in wb.sheetnames:
            ws_main = wb.create_sheet("Main")
            ws_main.append(self.headers)
        if "Log" not in wb.sheetnames:
            ws_log = wb.create_sheet("Log")
            ws_log.append(["Timestamp", "Action", "Updated Data"])
        wb.save(self.file_path)

    def load_data(self) -> None:
        """Load data from Excel file into session state"""
        try:
            if os.path.exists(self.file_path):
                df = pd.read_excel(self.file_path, sheet_name="Main", engine="openpyxl")
            else:
                # Create empty DataFrame with correct columns
                df = pd.DataFrame(columns=self.headers)
            st.session_state.df = df
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")
            st.session_state.df = pd.DataFrame(columns=self.headers)

    def create_grid(self) -> None:
        """Create and configure the editable data grid"""
        st.subheader("📝 Data Editor")

        # Ensure DataFrame exists with correct columns
        if st.session_state.df.empty:
            st.session_state.df = pd.DataFrame(columns=self.headers)

        # Create the editable grid
        edited_df = st.data_editor(
            st.session_state.df,
            num_rows="dynamic",
            use_container_width=True,
            key="data_editor",
            hide_index=True,
            column_config={
                "Account Number": st.column_config.TextColumn(
                    "Account Number",
                    width="medium",
                    required=True
                ),
                "Account Name": st.column_config.TextColumn(
                    "Account Name",
                    width="medium",
                    required=True
                ),
                **{
                    month: st.column_config.NumberColumn(
                        month,
                        width="small",
                        format="%.2f",
                        min_value=0,
                        required=False
                    )
                    for month in self.headers[2:]  # All month columns
                }
            }
        )

        if edited_df is not None:
            st.session_state.df = edited_df

    def add_new_row(self) -> None:
        """Add a new empty row to the dataframe"""
        if st.button("➕ Add Row"):
            empty_row = pd.DataFrame([[None] * len(self.headers)], columns=self.headers)
            st.session_state.df = pd.concat([st.session_state.df, empty_row], ignore_index=True)
            st.rerun()

    def save_data(self) -> None:
        """Save data to Excel file"""
        if st.button("💾 Save"):
            if self.validate_data_for_save():
                self.perform_save()
                st.success("✅ Data saved successfully!")

    def validate_data_for_save(self) -> bool:
        """Validate data before saving"""
        if st.session_state.df.empty:
            st.error("❌ Cannot save: The sheet has no data!")
            return False
        return True

    def perform_save(self) -> None:
        """Perform the save operation"""
        try:
            with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                st.session_state.df.to_excel(writer, sheet_name="Main", index=False)

                # Log the save action
                log_df = pd.DataFrame({
                    "Timestamp": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
                    "Action": ["Update"],
                    "Updated Data": [str(st.session_state.df.to_dict())]
                })

                # If Log sheet exists, read it and append
                try:
                    existing_log = pd.read_excel(self.file_path, sheet_name="Log")
                    log_df = pd.concat([existing_log, log_df], ignore_index=True)
                except:
                    pass

                log_df.to_excel(writer, sheet_name="Log", index=False)
        except Exception as e:
            st.error(f"Error saving data: {str(e)}")

    def clear_data(self) -> None:
        """Clear all data from the main sheet"""
        if st.button("🗑️ Clear Data"):
            if st.session_state.df.empty:
                st.warning("Sheet is already empty!")
                return

            st.session_state.df = pd.DataFrame(columns=self.headers)
            self.perform_save()
            st.success("🧹 Data cleared successfully!")

    def show_log(self) -> None:
        """Display the change log"""
        st.subheader("📋 Change Log")
        try:
            df_log = pd.read_excel(self.file_path, sheet_name="Log")
            st.dataframe(df_log, height=200)
        except Exception as e:
            st.error(f"Error loading log: {str(e)}")


def main():
    # Define headers
    headers = [
        "Account Number", "Account Name",
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]

    # Create an instance of ExcelApp
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