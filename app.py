import streamlit as st
import pandas as pd
from io import BytesIO
from merge_csv_only import process_files
from access_control_password import verify_user

st.set_page_config(page_title="Excel Merge Tool (CSV Only, Roles)", layout="wide")

# ---- SIDEBAR: LOGOUT & USER INFO ----
with st.sidebar:
    st.title("🔑 User Panel")
    if st.session_state.get("authenticated", False):
        st.success(f"Logged in as: {st.session_state.get('username', 'Unknown')}")
        st.info(f"Role: {st.session_state.get('role', '[none]')}")
        if st.button("🚪 Logout"):
            st.session_state['authenticated'] = False
            st.session_state['role'] = None
            st.session_state['username'] = ""
            st.session_state['merged_excel_bytes'] = None
            st.session_state['merged_df'] = None
            st.experimental_rerun()
    else:
        st.info("Please login to access the tool.")

st.markdown("<h1 style='text-align:center;'>🗂️ Excel Merge Tool</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align:center; color: gray;'>CSV Only + Roles</h4>", unsafe_allow_html=True)
st.markdown("---")

# ---- LOGIN FORM ----
if not st.session_state.get("authenticated", False):
    st.markdown("### 🔒 Please log in to continue")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
    if 'role' not in st.session_state:
        st.session_state['role'] = None
    if submitted:
        role = verify_user(username, password)
        if role:
            st.session_state['authenticated'] = True
            st.session_state['role'] = role
            st.session_state['username'] = username
            st.success(f"Logged in as {username} ({role})")
            st.experimental_rerun()
        else:
            st.error("Invalid username or password")
    st.stop()

role = st.session_state["role"]
username = st.session_state.get("username", "")

# ---- MAIN APP: UPLOAD & MERGE ----
st.header("1️⃣ Upload Files")
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
csv_files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)

if role in ["operator", "admin"]:
    st.header("2️⃣ Merge and Compare")
    if st.button("🔄 Start Merge"):
        if excel_file and csv_files:
            outputs = process_files(excel_file, csv_files)
            if outputs and isinstance(outputs, dict):
                st.success("✅ Merge complete! Download your file below:")
                merged_excel_bytes = outputs["detailed"].getvalue()
                st.download_button("📥 Download Merged Output (Excel)",
                                   data=merged_excel_bytes,
                                   file_name="merged_output.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.session_state["merged_excel_bytes"] = merged_excel_bytes
                merged_df = pd.read_excel(BytesIO(merged_excel_bytes), sheet_name="Merged Data", dtype=str)
                st.session_state["merged_df"] = merged_df
            else:
                st.error("❌ Merge failed.")
        else:
            st.warning("⚠️ Please upload all required files.")

elif role == "view":
    st.info("👁️ You have view-only access. Merge action is disabled.")

# Always retrieve merged_df from session_state
merged_df = st.session_state.get("merged_df")

if merged_df is not None:
    st.markdown("---")
    st.header("3️⃣ Field Comparison")
    cols = merged_df.columns.tolist()
    excel_cols = [col for col in cols if "_" not in col and col.lower() != "index" and col != "OVERRIDE ID"]
    csv_cols = [col for col in cols if "_" in col]
    csv_field_names = sorted(set(col.rsplit("_", 1)[0] for col in csv_cols), key=str.lower)

    col1, col2 = st.columns(2)
    with col1:
        sel_excel_col = st.selectbox("Select Excel Field", excel_cols, key="excel_field")
    with col2:
        sel_csv_field = st.selectbox("Select CSV Field Name", csv_field_names, key="csv_field_name")

    selected_csv_cols = [col for col in csv_cols if col.lower().startswith(sel_csv_field.lower() + "_")]
    display_cols = (
        ["OVERRIDE ID", sel_excel_col] + selected_csv_cols
        if "OVERRIDE ID" in merged_df.columns else
        [sel_excel_col] + selected_csv_cols
    )

    tabs = st.tabs(["Merged Data", "Field Comparison Viewer"])
    with tabs[0]:
        st.dataframe(merged_df, use_container_width=True, height=600)

    with tabs[1]:
        st.header("🔍 Field Comparison Viewer")
        if not selected_csv_cols:
            st.warning(f"No columns found for CSV field '{sel_csv_field}'. Check your merge or column names.")
        else:
            df_compare = merged_df[display_cols].copy()
            new_colnames = (
                ["OVERRIDE ID", "Excel"] + [col.split("_")[-1] for col in selected_csv_cols]
                if "OVERRIDE ID" in merged_df.columns else
                ["Excel"] + [col.split("_")[-1] for col in selected_csv_cols]
            )
            df_compare.columns = new_colnames

            csv_indices = list(range(2, len(new_colnames))) if "OVERRIDE ID" in merged_df.columns else list(range(1, len(new_colnames)))
            def highlight_diff(row):
                vals = [str(row.iloc[i]).strip() for i in csv_indices]
                consistent = len(set(vals)) <= 1
                highlights = ["" for _ in row.index]
                if not consistent:
                    for i in csv_indices:
                        highlights[i] = "background-color: #ffcccc"
                return highlights

            st.dataframe(
                df_compare.style.apply(highlight_diff, axis=1),
                use_container_width=True,
                height=600
            )
