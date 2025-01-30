import streamlit as st
import streamlit_authenticator as stauth
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime

# Pre-generated hashed passwords with emails
credentials = {
    "usernames": {
        "johndoe": {
            "name": "John Doe",
            "password": "$2b$12$WDhf5n2M1fe7R2rBBhn7e.jT0kE5umEx608XD7dLFUr55dB8zZk3a",
            "email": "johndoe@example.com"
        },
        "janesmith": {
            "name": "Jane Smith",
            "password": "$2b$12$iLipEEh.OarddPx5qXPw1eatZS4TV7CRR1ang2ZkBminzWa8TQNcy",
            "email": "janesmith@example.com"
        }
    }
}

# Initialize the authenticator
authenticator = stauth.Authenticate(
    credentials=credentials,
    cookie_name="app_cookie",
    key="random_key",
    cookie_expiry_days=1,
)

# Login
result = authenticator.login(location="main")
st.write("Login Function Output:", result)  # Debugging

if result:
    name, authentication_status, username = result
    if authentication_status:
        authenticator.logout("Logout", "sidebar")
        st.sidebar.write(f"Welcome, {name}!")

        st.title("Excel Data Entry & Viewer")

        # Input fields
        col_a = st.text_input("Enter value for Column A:")
        col_b = st.text_input("Enter value for Column B:")

        if st.button("Add to Excel"):
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

        st.subheader("Current Data in Excel:")
        df = pd.read_excel("example.xlsx", engine="openpyxl")
        st.dataframe(df)

    elif authentication_status == False:
        st.error("Username/password is incorrect")
    else:
        st.warning("Please enter your username and password")
else:
    st.error("Login function returned None. Check your credentials and library version.")
    st.stop()
