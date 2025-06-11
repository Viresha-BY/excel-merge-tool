import streamlit as st
import pandas as pd
from io import BytesIO
from merge_csv_only import process_files
from access_control_password import get_user_role

st.set_page_config(page_title="Excel Merge Tool (CSV Only, Roles)", layout="wide")
st.title("📘 Excel Merge Tool (CSV Only + Roles)")

token = get_user_role()
if token in [None, "none"]:
    st.stop()

st.markdown(f"**Role:** `{token}`")

if st.button("Logout"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.experimental_rerun()

st.header("Step 1: Upload Files")
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
csv_files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)

# Merge and store in session_state
if token in ["operator", "admin"]:
    st.header("Step 2: Merge and Compare")
    if st.button("🔄 Start Merge"):
        if excel_file and csv_files:
            outputs = process_files(excel_file, csv_files)
            if outputs and isinstance(outputs, dict):
                st.success("✅ Merge complete! Download file below:")
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

elif token == "view":
    st.info("👁️ You have view-only access. Merge action is disabled.")

# Always get merged_df from session_state so UI doesn't disappear on select
merged_df = st.session_state.get("merged_df")

# --- FIELD SELECTORS: always visible, above tabs ---
if merged_df is not None:
    cols = merged_df.columns.tolist()
    excel_cols = [col for col in cols if "_" not in col and col.lower() != "index" and col != "OVERRIDE ID"]
    csv_cols = [col for col in cols if "_" in col]
    csv_field_names = sorted(set(col.rsplit("_", 1)[0] for col in csv_cols), key=str.lower)

    st.subheader("Field Comparison Controls")
    col1, col2 = st.columns(2)
    with col1:
        sel_excel_col = st.selectbox("Select Excel Field", excel_cols, key="excel_field")
    with col2:
        sel_csv_field = st.selectbox("Select CSV Field Name", csv_field_names, key="csv_field_name")

    # Prepare which columns to display
    selected_csv_cols = [col for col in csv_cols if col.lower().startswith(sel_csv_field.lower() + "_")]

    display_cols = ["OVERRIDE ID", sel_excel_col] + selected_csv_cols if "OVERRIDE ID" in merged_df.columns else [sel_excel_col] + selected_csv_cols

    # Tab panes
    tabs = st.tabs(["Merged Data", "Field Comparison Viewer"])
    with tabs[0]:
        st.dataframe(merged_df, use_container_width=True, height=600)

    with tabs[1]:
        st.header("🔍 Field Comparison Viewer")
        if not selected_csv_cols:
            st.warning(f"No columns found for CSV field '{sel_csv_field}'. Check your merge or column names.")
        else:
            df_compare = merged_df[display_cols].copy()
            # Set pretty column labels
            new_colnames = ["OVERRIDE ID", "Excel"] + [col.split("_")[-1] for col in selected_csv_cols] if "OVERRIDE ID" in merged_df.columns else ["Excel"] + [col.split("_")[-1] for col in selected_csv_cols]
            df_compare.columns = new_colnames

            # Highlight if CSV values differ among themselves in a row (do not compare to Excel)
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