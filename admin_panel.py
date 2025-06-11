
import streamlit as st
import json
from pathlib import Path

CRED_FILE = "user_credentials.json"

def load_users():
    if Path(CRED_FILE).exists():
        with open(CRED_FILE, "r") as f:
            return json.load(f)
    return {}

def save_users(users):
    with open(CRED_FILE, "w") as f:
        json.dump(users, f, indent=4)

def user_management_ui():
    st.subheader("🔧 Manage Users")
    users = load_users()

    email_list = list(users.keys())
    selected = st.selectbox("Select user to edit", [""] + email_list)

    if selected:
        role = st.selectbox("Role", ["admin", "operator", "view"], index=["admin", "operator", "view"].index(users[selected]["role"]))
        password = st.text_input("New Password", type="password")
        if st.button("Update User"):
            users[selected]["role"] = role
            if password:
                users[selected]["password"] = password
            save_users(users)
            st.success(f"Updated {selected}")

        if st.button("Delete User"):
            del users[selected]
            save_users(users)
            st.warning(f"Deleted {selected}")

    st.markdown("---")
    st.subheader("➕ Add New User")
    new_email = st.text_input("Email").strip().lower()
    new_password = st.text_input("Password", type="password")
    new_role = st.selectbox("Assign Role", ["view", "operator", "admin"])
    if st.button("Add User"):
        if new_email in users:
            st.error("User already exists.")
        else:
            users[new_email] = {"password": new_password, "role": new_role}
            save_users(users)
            st.success(f"Added user {new_email}")
