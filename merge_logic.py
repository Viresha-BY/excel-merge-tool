# Full final version: CSV/JSON field-wise grouped + all summaries/unmatched/json sheets
import pandas as pd
import json
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from collections import defaultdict

red_fill = PatternFill(start_color="FFFF6666", end_color="FFFF6666", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")

def process_files(excel_file, csv_files, json_files):
    try:
        excel_df = pd.read_excel(excel_file)
        excel_df["MFL ID"] = excel_df["MFL ID"].astype(str).str.strip()
        excel_df["OVERRIDE ID"] = excel_df["OVERRIDE ID"].astype(str).str.strip()
        excel_df["DATE TIME PRE KO (UTC)"] = pd.to_datetime(excel_df["DATE TIME PRE KO (UTC)"], errors='coerce')
        excel_df["match_date"] = excel_df["DATE TIME PRE KO (UTC)"].dt.strftime("%Y-%m-%d")

        merged_df = excel_df.copy()
        csv_data = {}
        unmatched_data = {}
        summary_data = []
        consolidated_summary = {}

        for csv_file in csv_files:
            label = csv_file.name.split('.')[0]
            df = pd.read_csv(csv_file)
            df["clientContentId"] = df["clientContentId"].astype(str).str.strip()
            df.set_index("clientContentId", inplace=True, drop=False)
            csv_data[label] = df
            unmatched_data[label] = df.copy()

        for label, df in csv_data.items():
            for col in df.columns:
                if col not in ("competitionId", "Day", "launchPeriod", "rightsId", "Source"):
                    merged_df[f"{col}_{label}"] = None

        for i in merged_df.index:
            mfl_id = merged_df.at[i, "MFL ID"]
            for label, df in csv_data.items():
                if mfl_id in df.index:
                    row = df.loc[mfl_id]
                    if isinstance(row, pd.DataFrame):
                        row = row.iloc[0]
                    for col in df.columns:
                        if col not in ("competitionId", "Day", "launchPeriod", "rightsId", "Source"):
                            merged_df.at[i, f"{col}_{label}"] = row.get(col, None)
                    unmatched_data[label].drop(mfl_id, inplace=True, errors='ignore')

        # Add JSON
        json_data_by_label = {}
        for json_file in json_files:
            label = json_file.name.split('.')[0]
            json_data = json.load(json_file)
            records = []
            for obj in json_data:
                e = obj["event"]
                oa_id = e.get("oaId", "")
                bcast = e.get("broadcasts", {}).get(oa_id, {})
                records.append({
                    "overrideId": e.get("overrideId", [{}])[0].get("id"),
                    "date": e.get("streamStartTime", "")[:10],
                    f"oaId_{label}": oa_id,
                    f"streamStartTime_{label}": e.get("streamStartTime", ""),
                    f"streamEndTime_{label}": e.get("streamEndTime", ""),
                    f"heEventTypeName_{label}": e.get("heEventTypeName", ""),
                    f"drmRequired_{label}": e.get("drmRequired", ""),
                    f"regions_{label}": ", ".join(e.get("regions", [])) if isinstance(e.get("regions", []), list) else e.get("regions", ""),
                    f"outputSuppressionMode_{label}": bcast.get("outputSuppressionMode", ""),
                    f"assetName_{label}": bcast.get("name", ""),
                    f"template_{label}": bcast.get("template", ""),
                    f"heResilience_{label}": e.get("heResilience", ""),
                    f"competitionId_{label}": e.get("competitionId", ""),
                    f"closedCaptioning_{label}": ", ".join(e.get("closedCaptioning", [])) if isinstance(e.get("closedCaptioning", []), list) else e.get("closedCaptioning", ""),
                    f"description_{label}": e.get("description", "")
                })
            json_df = pd.DataFrame(records).set_index(["overrideId", "date"])
            json_data_by_label[label] = json_df

            for field in json_df.columns:
                if field not in merged_df.columns:
                    merged_df[field] = None

            for i in merged_df.index:
                key = (str(merged_df.at[i, "OVERRIDE ID"]).strip(), merged_df.at[i, "match_date"])
                if key in json_df.index:
                    row = json_df.loc[key]
                    if isinstance(row, pd.DataFrame):
                        row = row.iloc[0]
                    for field in json_df.columns:
                        merged_df.at[i, field] = row.get(field, None)

        merged_df.drop(columns=["match_date"], inplace=True)

        # ✅ Final column grouping by field name across labels
        all_labels = list(csv_data.keys()) + list(json_data_by_label.keys())
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

        headers = [cell.value for cell in ws1[1]]
        for col_idx, col_name in enumerate(headers, start=1):
            if col_name.startswith("clientContentId_"):
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

        # ✅ Summary
        ws2 = wb.create_sheet("Summary")
        ws2.append(["File", "Type", "Total Rows", "Matched to Excel", "Extra", "Missing"])
        for label, df in csv_data.items():
            excel_ids = set(excel_df["MFL ID"])
            csv_ids = set(df["clientContentId"])
            matched = len(excel_ids & csv_ids)
            extra = len(csv_ids - excel_ids)
            missing = len(excel_ids - csv_ids)
            ws2.append([label, "CSV", len(csv_ids), matched, extra, missing])
        used_keys = set(zip(excel_df["OVERRIDE ID"], excel_df["DATE TIME PRE KO (UTC)"].dt.strftime("%Y-%m-%d")))
        for label, json_df in json_data_by_label.items():
            filtered = json_df.loc[json_df.index.isin(used_keys)]
            total = filtered.shape[0]
            filled = sum(filtered[field].notna().sum() for field in filtered.columns)
            empty = sum(filtered[field].isna().sum() for field in filtered.columns)
            ws2.append([label, "JSON", total, total, "-", empty])

        # ✅ Consolidated Summary
        ws3 = wb.create_sheet("Consolidated Summary")
        ws3.append(["Source", "File", "Field", "Matched", "Missing", "Mismatched"])
        for label, df in csv_data.items():
            for col in df.columns:
                if col not in ("competitionId", "Day", "launchPeriod", "rightsId", "Source"):
                    match, missing, mismatch = 0, 0, 0
                    for i in merged_df.index:
                        value = merged_df.at[i, f"{col}_{label}"]
                        if value is None or value == "":
                            missing += 1
                        elif col == "clientContentId" and value != merged_df.at[i, "MFL ID"]:
                            mismatch += 1
                        else:
                            match += 1
                    ws3.append(["CSV", label, col, match, missing, mismatch])
        for label, json_df in json_data_by_label.items():
            filtered = json_df.loc[json_df.index.isin(used_keys)]
            for field in json_df.columns:
                total = filtered.shape[0]
                non_empty = filtered[field].notna().sum()
                missing = total - non_empty
                ws3.append(["JSON", label, field, non_empty, missing, "-"])

        # ✅ Unmatched Sheets
        for label, df in unmatched_data.items():
            ws = wb.create_sheet(f"Unmatched_{label}")
            for r in dataframe_to_rows(df.reset_index(drop=True), index=False, header=True):
                ws.append(r)

        # ✅ JSON Sheets
        for label, json_df in json_data_by_label.items():
            df_reset = json_df.reset_index()
            filtered = df_reset[df_reset.apply(lambda x: (str(x["overrideId"]).strip(), str(x["date"]).strip()) in used_keys, axis=1)]
            ws = wb.create_sheet(f"JSON_{label}")
            for r in dataframe_to_rows(filtered, index=False, header=True):
                ws.append(r)

        wb.save(output)
        output.seek(0)
        return output
    except Exception as e:
        import traceback
        traceback.print_exc()
        return None

