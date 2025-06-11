import streamlit as st
import json

def get_user_role():
    user_email = st.text_input("Enter your email").strip().lower()
    password = st.text_input("Enter password", type="password")
    
    try:
        with open("user_credentials.json", "r") as f:
            users = json.load(f)
        user = users.get(user_email)
        if user and user.get("password") == password:
            st.success(f"Welcome {user_email} ({user['role']})")
            return user["role"]
        else:
            st.error("Invalid credentials")
            return "none"
    except Exception as e:
        st.error("Credential file error.")
        return "none"
