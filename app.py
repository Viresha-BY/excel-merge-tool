import streamlit as st
import pandas as pd
import requests
import json
from io import BytesIO
from merge_logic import process_files

st.set_page_config(page_title="Excel Merge Tool with API", layout="wide")
st.title("📘 Excel Merge Tool (CSV + JSON + API)")

excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
csv_files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)

# JSON source toggle
json_mode = st.radio("Select JSON Source", ["Upload JSON Files", "Fetch from Live API"])

json_files = []

if json_mode == "Upload JSON Files":
    uploaded_jsons = st.file_uploader("Upload JSON Files", type=["json"], accept_multiple_files=True)
    if uploaded_jsons:
        json_files = uploaded_jsons

elif json_mode == "Fetch from Live API":
    st.markdown("Fetching data from predefined live API endpoints...")
    api_urls = {
        "DCG": "https://cluster-1.uksouth-3.streaming.mediakind.com/dazn/events?",
        "DCH": "https://cluster-1.westeurope-3.streaming.mediakind.com/dazn/events"
    }

    for label, url in api_urls.items():
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                data = response.json()
                json_files.append(BytesIO(json.dumps(data).encode("utf-8")))
                json_files[-1].name = f"{label}.json"
                st.success(f"✅ Fetched {label} ({len(data)} records)")
            else:
                st.warning(f"⚠️ Failed to fetch {label} - Status Code {response.status_code}")
        except Exception as e:
            st.error(f"❌ Error fetching {label}: {e}")

if st.button("🔄 Start Merge") and excel_file and csv_files and json_files:
    with st.spinner("Processing... Please wait."):
        output = process_files(excel_file, csv_files, json_files)
        if output:
            st.success("✅ Merge Complete. Download your files below:")
            st.download_button("📥 Download Output File 1", output.getvalue(), "merged_output.xlsx")
        else:
            st.error("❌ Processing failed. Please check inputs.")
