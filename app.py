import streamlit as st
import pandas as pd
from io import BytesIO
from merge_csv_only import process_files
from access_control_password import verify_user
from validation_logic import style_dataframe

st.set_page_config(page_title="Excel & CSV files Merge Comparison Tool", layout="wide")

# ---- Columns to hide (case-insensitive substring match) ----
HIDE_SUBSTRS = [
    "description", "languagemappingname", "source", "eventurl", "cancelled", "rightsId",
    "streamstartdatetime", "streamenddatetime", "tier", "version", "launchperiod",
]

def hide_col(col):
    col_lower = col.lower()
    return any(substr in col_lower for substr in HIDE_SUBSTRS)

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
            st.rerun()
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
            st.rerun()
        else:
            st.error("Invalid username or password")
    st.stop()

role = st.session_state["role"]
username = st.session_state.get("username", "")

# ---- MAIN APP: UPLOADS ----
st.markdown("## <span style='color:#1a73e8;font-weight:bold;'>⬆️ Upload PowerBI Export file</span>", unsafe_allow_html=True)
excel_file = st.file_uploader("", type=["xlsx"], key="excel_file")

st.markdown("## <span style='color:#34a853;font-weight:bold;'>⬆️ Upload All DynamoDB files</span>", unsafe_allow_html=True)
csv_files = st.file_uploader(
    "",
    type=["csv"],
    accept_multiple_files=True,
    key="csv_files"
)

# ---- MERGE LOGIC ----
if role in ["operator", "admin"]:
    st.header("2️⃣ Merge and Compare")
    if st.button("🔄 Start Merge"):
        if excel_file and csv_files:
            outputs = process_files(excel_file, csv_files)
            if outputs and isinstance(outputs, dict):
                st.success("✅ Merge complete! Download your files below:")
                merged_excel_bytes = outputs["detailed"].getvalue()
                st.session_state["merged_excel_bytes"] = merged_excel_bytes
                merged_df = pd.read_excel(
                    BytesIO(merged_excel_bytes), 
                    sheet_name="Merged Data", 
                    dtype=str, keep_default_na=False
                )
                merged_df = merged_df.fillna("")
                st.session_state["merged_df"] = merged_df
            else:
                st.error("❌ Merge failed.")
        else:
            st.warning("⚠️ Please upload all required files.")

elif role == "view":
    st.info("👁️ You have view-only access. Merge action is disabled.")

merged_df = st.session_state.get("merged_df")

if merged_df is not None:
    st.markdown("---")
    st.header("3️⃣ Download Options & Field Comparison")

    # --- UI Column Filtering ---
    ui_cols = [col for col in merged_df.columns if not hide_col(col)]

    # --- Download buttons ---
    col_dl1, col_dl2, col_dl3 = st.columns(3)
    with col_dl1:
        st.download_button(
            "📥 Download Merged Output (raw, all columns)",
            data=st.session_state["merged_excel_bytes"],
            file_name="merged_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_dl2:
        styled_validated = style_dataframe(merged_df)
        output_validated = BytesIO()
        with pd.ExcelWriter(output_validated, engine="openpyxl") as writer:
            styled_validated.to_excel(writer, index=False, sheet_name="StyledData")
        output_validated.seek(0)
        st.download_button(
            "📥 Download Validated Output (all columns, with colors)",
            data=output_validated,
            file_name="validated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    tabs = st.tabs(["Merged Data", "Field Comparison Viewer"])

    # --- Tab 1: Merged Data (UI, hidden columns) ---
    with tabs[0]:
        styled_ui = style_dataframe(merged_df[ui_cols])
        st.dataframe(styled_ui, use_container_width=True, height=600)

    # --- Tab 2: Field Comparison (hidden columns) ---
    with tabs[1]:
        st.header("🔍 Field Comparison Viewer")
        excel_cols = [col for col in ui_cols if "_" not in col and col.lower() != "index" and col != "OVERRIDE ID"]
        csv_cols = [col for col in ui_cols if "_" in col]
        csv_field_names = sorted(set(col.rsplit("_", 1)[0] for col in csv_cols), key=str.lower)

        col1, col2 = st.columns(2)
        with col1:
            sel_excel_col = st.selectbox("Select Excel Field", excel_cols, key="excel_field")
        with col2:
            sel_csv_field = st.selectbox("Select CSV Field Name", csv_field_names, key="csv_field_name")

        selected_csv_cols = [col for col in csv_cols if col.lower().startswith(sel_csv_field.lower() + "_")]
        compare_cols = (
            (["OVERRIDE ID"] if "OVERRIDE ID" in ui_cols else []) +
            [sel_excel_col] + selected_csv_cols
        )
        compare_cols = [col for col in compare_cols if col in ui_cols]

        if not selected_csv_cols:
            st.warning(f"No columns found for CSV field '{sel_csv_field}'. Check your merge or column names.")
        else:
            subset_df = merged_df[compare_cols]
            styled_subset = style_dataframe(subset_df)
            st.dataframe(styled_subset, use_container_width=True, height=600)

            # --- Download comparison fields (with colors) ---
            output_compare = BytesIO()
            with pd.ExcelWriter(output_compare, engine="openpyxl") as writer:
                styled_subset.to_excel(writer, index=False, sheet_name="FieldComparison")
            output_compare.seek(0)
            st.download_button(
                "📥 Download Selected Fields (Field Comparison, with colors)",
                data=output_compare,
                file_name="selected_fields.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )