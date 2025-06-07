# 📘 Excel Merge Tool (CSV + JSON + Summary)

A Streamlit-based web tool to merge and validate Excel data with multiple CSVs and JSON files.

## 🚀 Getting Started

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
streamlit run app.py
```

### 3. Upload Files
- Excel file (XLSX) with `MFL ID`, `OVERRIDE ID`, `DATE TIME PRE KO (UTC)`
- CSV files with `clientContentId`
- JSON files containing event metadata (joined via overrideId + date)

## 📂 Project Structure
```
excel-merge-tool/
├── app.py
├── merge_logic.py
├── requirements.txt
├── README.md
```

## ✅ Features
- Excel + multi-CSV match
- Multi-JSON enrichment (DCG, DCH, etc.)
- Highlighting mismatches + missing values
- Summary + consolidated report sheets
