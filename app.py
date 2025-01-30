import streamlit as st
import streamlit_authenticator as stauth
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime

# Authentication configuration
names = ["John Doe", "Jane Smith"]
usernames = ["johndoe", "janesmith"]
passwords = ["password123", "password456"]  # Use hashed passwords in production

# Hash the passwords for security (only run this once, not every time)
hashed_passwords = stauth.Hasher(passwords).generate()

authenticator = stauth.Authenticate(
    names, usernames, hashed_passwords,
    "app_cookie", "random_key", cookie_expiry_days=1
)

# Login
name, authentication_status, username = authenticator.login("Login", "main")

if authentication_status:
    # Show the main app if login is successful
    authenticator.logout("Logout", "sidebar")
    st.sidebar.write(f"Welcome, {name}!")

    # App content
    st.title("Excel Data Entry & Viewer")

    # Input fields
    col_a = st.text_input("Enter value for Column A:")
    col_b = st.text_input("Enter value for Column B:")

    if st.button("Add to Excel"):
        # Add data to Excel file
        file_path = "example.xlsx"
        wb = load_workbook(file_path)
        ws = wb.active
        next_row = ws.max_row + 1
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws[f"A{next_row}"] = col_a
        ws[f"B{next_row}"] = col_b
        ws[f"C{next_row}"] = timestamp
        wb.save(file_path)
        st.success("Data added successfully!")

    # Display current data
    st.subheader("Current Data in Excel:")
    df = pd.read_excel("example.xlsx", engine="openpyxl")
    st.dataframe(df)

elif authentication_status == False:
    st.error("Username/password is incorrect")
elif authentication_status == None:
    st.warning("Please enter your username and password")
