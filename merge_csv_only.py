import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import defaultdict

red_fill = PatternFill(start_color="FFFF6666", end_color="FFFF6666", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
orange_fill = PatternFill(start_color="FFFF9900", end_color="FFFF9900", fill_type="solid")

def _is_invalid_key(val):
    try:
        return pd.isna(val) or str(val).strip().lower() in ("", "nan", "none")
    except Exception:
        return True

def clean_override_id(val):
    try:
        f = float(val)
        i = int(f)
        return str(i)
    except Exception:
        return str(val).strip()

def process_files(excel_file, csv_files):
    try:
        excel_df = pd.read_excel(excel_file)
        excel_df["MFL ID"] = excel_df["MFL ID"].apply(clean_override_id)
        excel_df["OVERRIDE ID"] = excel_df["OVERRIDE ID"].apply(clean_override_id)
        excel_df["DATE TIME PRE KO (UTC)"] = pd.to_datetime(excel_df["DATE TIME PRE KO (UTC)"], errors='coerce', dayfirst=True)
        excel_df["match_date"] = excel_df["DATE TIME PRE KO (UTC)"].dt.strftime("%Y-%m-%d")

        merged_df = excel_df.copy()
        csv_data = {}
        unmatched_data = {}

        for csv_file in csv_files:
            label = csv_file.name.split('.')[0]
            df = pd.read_csv(csv_file, dtype=str)
            df["clientContentId"] = df["clientContentId"].apply(clean_override_id)
            df.set_index("clientContentId", inplace=True, drop=False)
            csv_data[label] = df
            unmatched_data[label] = df.copy()

        # Batch-add all CSV columns for each label
        all_new_cols = []
        for label, df in csv_data.items():
            for col in df.columns:
                if col not in ("competitionId", "Day", "launchPeriod", "rightsId", "Source"):
                    all_new_cols.append(f"{col}_{label}")
        new_cols_to_add = [c for c in all_new_cols if c not in merged_df.columns]
        if new_cols_to_add:
            merged_df = pd.concat([merged_df, pd.DataFrame({col: None for col in new_cols_to_add}, index=merged_df.index)], axis=1)

        for i in merged_df.index:
            mfl_id = merged_df.at[i, "MFL ID"]
            if _is_invalid_key(mfl_id):
                continue
            for label, df in csv_data.items():
                if mfl_id in df.index:
                    row = df.loc[mfl_id]
                    if isinstance(row, pd.DataFrame):
                        row = row.iloc[0]
                    for col in df.columns:
                        if col not in ("competitionId", "Day", "launchPeriod", "rightsId", "Source"):
                            val = str(row.get(col, "")) if pd.notna(row.get(col, "")) else ""
                            merged_df.at[i, f"{col}_{label}"] = val
                    unmatched_data[label].drop(mfl_id, inplace=True, errors='ignore')

        merged_df.drop(columns=["match_date"], inplace=True)

        # Final column grouping by field name across labels
        all_labels = list(csv_data.keys())
        excel_cols = [col for col in merged_df.columns if "_" not in col]
        grouped_cols = [col for col in merged_df.columns if "_" in col]

        field_groups = defaultdict(list)
        for col in grouped_cols:
            base, label = col.rsplit("_", 1)
            field_groups[base].append(col)

        reordered_cols = excel_cols + [col for base in sorted(field_groups) for col in sorted(field_groups[base])]
        merged_df = merged_df[reordered_cols]

        output = BytesIO()
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Merged Data"
        for r in dataframe_to_rows(merged_df, index=False, header=True):
            ws1.append(r)

        # Highlight mismatches between Excel and each CSV field
        headers = [cell.value for cell in ws1[1]]
        for col_idx, col_name in enumerate(headers, start=1):
            if "_" in col_name and not col_name.startswith("clientContentId_"):
                # Compare with base Excel column
                base_col = col_name.split("_")[0]
                if base_col in headers:
                    base_col_idx = headers.index(base_col) + 1
                    for row_idx in range(2, ws1.max_row + 1):
                        val = ws1.cell(row=row_idx, column=col_idx).value
                        base_val = ws1.cell(row=row_idx, column=base_col_idx).value
                        if str(val).strip() != str(base_val).strip():
                            ws1.cell(row=row_idx, column=col_idx).fill = orange_fill

            elif col_name.startswith("clientContentId_"):
                for row_idx in range(2, ws1.max_row + 1):
                    val = ws1.cell(row=row_idx, column=col_idx).value
                    mfl_val = ws1.cell(row=row_idx, column=headers.index("MFL ID") + 1).value
                    if val and mfl_val and str(val).strip() != str(mfl_val).strip():
                        ws1.cell(row=row_idx, column=col_idx).fill = red_fill

            elif "_" in col_name:
                for row_idx in range(2, ws1.max_row + 1):
                    val = ws1.cell(row=row_idx, column=col_idx).value
                    if val is None or str(val).strip() == "":
                        ws1.cell(row=row_idx, column=col_idx).fill = yellow_fill

        # Add unmatched CSV rows as separate sheets
        for label, unmatched in unmatched_data.items():
            ws_csv = wb.create_sheet(f"Unmatched_{label[:25]}")
            if not unmatched.empty:
                for r in dataframe_to_rows(unmatched, index=False, header=True):
                    ws_csv.append(r)

        # Only one output file now!
        wb.save(output)
        output.seek(0)
        return {"detailed": output}

    except Exception as e:
        print("Error in process_files:", e)
        return None

def merge_files(excel_file, csv_files):
    output = process_files(excel_file, csv_files)
    return output, None, None

def save_output(merged_df, mismatch_summary, json_diff, filename):
    if merged_df is not None and hasattr(merged_df, "to_excel"):
        merged_df.to_excel(filename, index=False)
    pass