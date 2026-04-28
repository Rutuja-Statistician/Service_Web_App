import io
import sys
import time
import pandas as pd
import streamlit as st
from datetime import datetime
from streamlit_gsheets import GSheetsConnection
import gspread
from google.oauth2.service_account import Credentials

# To create connection
def get_gsheet_conn():
    """Create authenticated gspread connection using secrets."""
    creds_dict = {
        "type": st.secrets["connections"]["gsheets"]["type"],
        "project_id": st.secrets["connections"]["gsheets"]["project_id"],
        "private_key_id": st.secrets["connections"]["gsheets"]["private_key_id"],
        "private_key": st.secrets["connections"]["gsheets"]["private_key"],
        "client_email": st.secrets["connections"]["gsheets"]["client_email"],
        "client_id": st.secrets["connections"]["gsheets"]["client_id"],
        "auth_uri": st.secrets["connections"]["gsheets"]["auth_uri"],
        "token_uri": st.secrets["connections"]["gsheets"]["token_uri"],
    }

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client

def connect_gsheet():
    try:
        # Google sheet Connection  
        client = get_gsheet_conn()
        SPREADSHEET_ID = "1XlKbbbdJ3ySHwDxm_liTTOQf1QhliJHsFFdIz_r-CDY"
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        print("Connection successful...!!!")
        return spreadsheet
    except Exception as e:
        print(f"Unable to connect google sheet: {e}")
        show_popup(f"Unable to connect google sheet: {e}", type="error")
        

def show_popup(message, type = "success"):
    if type == "success":
        st.toast(f"✅ {message}")
    elif type == "error" :
        st.toast(f"❌ {message}")
    elif type == "warning":
        st.toast(f"⚠️ {message}")
    elif type == "info":
        st.toast(f"ℹ️ {message}")

def func1(raw_file):
    try:
        spreadsheet = connect_gsheet()

        # Open or create the worksheet
        try:
            detailData_worksheet = spreadsheet.worksheet("Detailed_Data")
        except gspread.WorksheetNotFound:
            detailData_worksheet = spreadsheet.add_worksheet("Detailed_Data", rows=5000, cols=30)
            
        # Clear existing data and write fresh
        detailData_worksheet.clear()

        data = pd.read_excel(raw_file)
        data.columns = data.columns.str.lower().str.replace(" ","_").str.replace(".", "_").str.strip()
        # To select the subset of the dataframe from the complete data
        selected_columns = ["service_id","circle", "customer_type", "call_date", "updatedate", "status_code"]
        data = data[selected_columns]
        data["service_id"] = data["service_id"].astype(str)
        data["call_date"] = pd.to_datetime(data["call_date"]).dt.normalize()
        data["updatedate"] = pd.to_datetime(data["updatedate"]).dt.normalize()

        todayDate = pd.to_datetime('today').date()
        data = data[data["call_date"].dt.date != todayDate]
        data = data[data["circle"].str.lower().str.strip() != "india"]

        data["today_date"] = pd.to_datetime(todayDate)
        data["age_from_call_reg"] = data["today_date"] - data["call_date"]
        data["age_from_call_update"] = data["today_date"] - data["updatedate"]

        # status_data = pd.read_excel(statuswise_file)
        norms_worksheet = spreadsheet.worksheet("Norms_Data")
        norms_data = norms_worksheet.get_all_records()
        status_data = pd.DataFrame(norms_data)

        status_data.columns = status_data.columns.str.lower().str.strip().str.replace(" ", "_")

        merged_data = data.merge(status_data[["status","team", "number"]], left_on= "status_code", right_on="status", how= "left")
        merged_data.to_excel("Raw_Data.xlsx", index= False)
        
        # Adding filter on teams, choosing customer xperience
        merged_data = merged_data[merged_data["team"].str.lower().str.strip() == "customer xperience"]

        merged_data["age_reg_days"] = merged_data["age_from_call_reg"].dt.days 
        merged_data["age_update_days"] = merged_data["age_from_call_update"].dt.days 

        def assign_category(row):
            status = str(row["status"]).strip().lower()
            num = row["number"]
            if pd.isna(num): return ""
            age = row["age_reg_days"] if status in ["open", "work_allocated"] else row["age_update_days"]
            if pd.isna(age): return ""

            if age > num: return "Red Call"
            elif age == num: return "Encroaching1"
            elif age == num - 1: return "Encroaching2"
            return ""

        merged_data["category"] = merged_data.apply(assign_category, axis=1)
        merged_data["red_call_flag"] = (merged_data["category"] == "Red Call").astype(int)
        merged_data["enc1_flag"] = (merged_data["category"] == "Encroaching1").astype(int)
        merged_data["enc2_flag"] = (merged_data["category"] == "Encroaching2").astype(int)
        
        # To write data in google sheet
        if merged_data is not None and not merged_data.empty:
            
            # Convert DataFrame to list of lists
            data_to_write = [merged_data.columns.tolist()] + merged_data.astype(str).values.tolist()
            detailData_worksheet.update(data_to_write)
            
            show_popup("Data stored in the database", type = "success")
        else:
            show_popup("No data found after filtering...!", type = "info")
        return merged_data

    except Exception as e:
        print(f"Error in func1: {e}")
        show_popup(f"Error in function is: {e}", type= "error")

def circlewise_platter(merged_data):
    try:
        spreadsheet = connect_gsheet()
        summary = merged_data.groupby("circle").agg({
            "red_call_flag": "sum", "enc1_flag": "sum", "enc2_flag": "sum"
        }).reset_index()

        summary = summary.rename(columns={
            "circle": "Circle", "red_call_flag": "Red Call",
            "enc1_flag": "Encroaching1", "enc2_flag": "Encroaching2"
        })

        summary["Platter1"] = summary["Red Call"] + summary["Encroaching1"]
        summary["Platter2"] = summary["Platter1"] + summary["Encroaching2"]

        col_order = ["Circle", "Red Call", "Encroaching1", "Platter1", "Encroaching2", "Platter2"]
        summary = summary[col_order]

        # Sort by Platter1 descending BEFORE adding Total row
        summary = summary.sort_values("Platter1", ascending=False).reset_index(drop=True)

        totals = summary.select_dtypes(include='number').sum()
        total_row = pd.DataFrame([totals])
        total_row["Circle"] = "Total"
        summary = pd.concat([summary, total_row], ignore_index=True)

        # Write to Google Sheet
        try:
            worksheet = spreadsheet.worksheet("Daily Circlewise Platter")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("Daily Circlewise Platter", rows=5000, cols=30)

        worksheet.clear()
        data_to_write = [summary.columns.tolist()] + summary.astype(str).values.tolist()
        worksheet.update(data_to_write)

        # Call tracker with circlewise summary
        tracker(summary)

        # st.subheader("Circlewise Platter Preview")
        # st.dataframe(summary, use_container_width=True)

        show_popup("Circlewise platter report created successfully!", type="success")

    except Exception as e:
        print(f"Error in circlewise platter function: {e}")
        show_popup(f"Error in circlewise_platter is : {e}", type="error") 


def tracker(df):
    try:
        spreadsheet = connect_gsheet()
        today_str = datetime.now().strftime("%Y-%m-%d")
        time_str  = datetime.now().strftime("%H:%M")
        print("The time string is:",time_str)

        # --- Get or create Tracker worksheet ---
        try:
            worksheet = spreadsheet.worksheet("Tracker")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("Tracker", rows=1000, cols=50)

        # --- Read existing data (returns list of lists) ---
        existing_data = worksheet.get_all_values()

        # --- Check if first DATA row has today's date (row index 1, skip header) ---
        if not existing_data or len(existing_data) < 2 or existing_data[1][0] != today_str:
            worksheet.clear()
            existing_data = []

        # --- Extract Circle + Platter1 + Platter2 ---
        platter_data = df[["Circle", "Platter1", "Platter2"]].copy()
        platter_data["Circle"] = platter_data["Circle"].astype(str)
        platter_data["Platter1"] = pd.to_numeric(platter_data["Platter1"], errors= 'coerce').fillna(0)
        platter_data["Platter2"] = pd.to_numeric(platter_data["Platter2"], errors= 'coerce').fillna(0)

        new_p1_col = f"{time_str} P1"
        new_p2_col = f"{time_str} P2"

        # Helper: push Total row(s) to the very bottom
        def sort_total_to_bottom(df_in):
            is_total = df_in["Circle"].str.strip().str.lower() == "total"
            return pd.concat(
                [df_in[~is_total], df_in[is_total]],
                ignore_index=True
            )

        if not existing_data:
            # --- First run of the day: write fresh ---
            print("No existing data, writing fresh...")
            header = ["Date", "Circle", new_p1_col, new_p2_col]
            rows = [header]
            for _, row in platter_data.iterrows():
                rows.append([today_str, str(row["Circle"]), str(row["Platter1"]), str(row["Platter2"])])
            worksheet.update(rows)

        else:
            # --- Subsequent runs: append new P1/P2 columns ---
            print("Existing data found, appending new columns...")
            print("Existing headers:", existing_data[0])

            # Convert list of lists → DataFrame
            existing_df = pd.DataFrame(existing_data[1:], columns=existing_data[0])
            existing_df["Circle"] = existing_df["Circle"].astype(str)

            # Build lookup dict: Circle → (P1, P2)
            platter_dict = {
                str(row["Circle"]): (str(row["Platter1"]), str(row["Platter2"]))
                for _, row in platter_data.iterrows()
            }

            # Find circles in new data that are NOT in existing sheet → add as new rows
            existing_circles = set(existing_df["Circle"].tolist())
            new_circles = [c for c in platter_dict if c not in existing_circles]

            if new_circles:
                print(f"New circles found, adding rows: {new_circles}")
                empty_row = {col: "0" for col in existing_df.columns}
                new_rows = []
                for circle in new_circles:
                    row = empty_row.copy()
                    row["Date"]   = today_str
                    row["Circle"] = circle
                    new_rows.append(row)
                new_rows_df = pd.DataFrame(new_rows)
                existing_df = pd.concat([existing_df, new_rows_df], ignore_index=True)

            # Add new time-stamped columns for ALL rows (existing + newly added)
            existing_df[new_p1_col] = existing_df["Circle"].map(
                lambda c: platter_dict.get(c, ("0", "0"))[0]
            )
            existing_df[new_p2_col] = existing_df["Circle"].map(
                lambda c: platter_dict.get(c, ("0", "0"))[1]
            )

            # Push Total row to the bottom before writing
            existing_df = sort_total_to_bottom(existing_df)

            # Write back updated data
            updated_rows = [existing_df.columns.tolist()] + existing_df.values.tolist()
            # middle_rows = sorted(updated_rows[1:len(updated_rows)-1],key = lambda x: x[1])

            # Find the index of the latest P1 column (last P1 in headers)
            headers = updated_rows[0]
            p1_indices = [i for i, h in enumerate(headers) if h.endswith("P1")]
            latest_p1_idx = p1_indices[-1]  # ✅ Last P1 column = most recent time

            # Sort middle rows by latest P1 descending, keep header and Total row in place
            middle_rows = sorted(
                updated_rows[1:len(updated_rows)-1],
                key=lambda x: int(float(x[latest_p1_idx])) if x[latest_p1_idx] not in ("", None) else 0,
                reverse=True 
            )

            new_rows = [headers] + middle_rows + [updated_rows[-1]]  # header + sorted + Total


            # new_rows = [updated_rows[0]] + middle_rows + [updated_rows[-1]]
            worksheet.clear()
            worksheet.update(new_rows)

        show_popup("Tracker updated successfully!", type="success")

    except Exception as e:
        print(f"Error in tracker function is: {e}")
        show_popup(f"Error in tracker function is: {e}", type="error")

def statuswise_platter(merged_data):
    try:
        spreadsheet = connect_gsheet()

        summary = merged_data.groupby("status_code").agg({
            "red_call_flag": "sum", "enc1_flag": "sum", "enc2_flag": "sum"
        }).reset_index()

        summary = summary.rename(columns={
            "status_code": "Status", "red_call_flag": "Red Call",
            "enc1_flag": "Encroaching1", "enc2_flag": "Encroaching2"
        })

        summary["Platter1"] = summary["Red Call"] + summary["Encroaching1"]
        summary["Platter2"] = summary["Platter1"] + summary["Encroaching2"]
        
        col_order = ["Status", "Red Call", "Encroaching1", "Platter1","Encroaching2", "Platter2"]
        summary = summary[col_order]

        # Sort by Platter1 descending BEFORE adding Total row
        summary = summary.sort_values("Platter1", ascending=False).reset_index(drop=True)
        
        totals = summary.select_dtypes(include='number').sum()
        total_row = pd.DataFrame([totals])
        total_row["Status"] = "Total"
        summary = pd.concat([summary, total_row], ignore_index=True)

        # To write data in google sheet
        if summary is not None and not summary.empty:
            # Open or create the worksheet
            try:
                worksheet = spreadsheet.worksheet("Statuswise Platter")
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet("Statuswise Platter", rows=5000, cols=30)
            
            # Clear existing data and write fresh
            worksheet.clear()
            
            # Convert DataFrame to list of lists
            data_to_write = [summary.columns.tolist()] + summary.astype(str).values.tolist()
            worksheet.update(data_to_write)
        
        show_popup("Statuswise platter report created successfully!", type = "success")        
    except Exception as e:
        print(f"Error in statuswise_platter function : {e}")
        show_popup(f"Error in statuswise_platter is : {e}", type = "error")

def billing_code_status_platter(merged_data):
    try:
        spreadsheet = connect_gsheet()

        # To write data in google sheet
        try:
            worksheet = spreadsheet.worksheet("Billing Code")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("Billing Code", rows=5000, cols=30)
            
        # Clear existing data and write fresh
        worksheet.clear()

        summary = merged_data[merged_data["status_code"].str.upper().str.strip() == "BILLING_CODE_PROBLEM"]
        
        if not merged_data.empty:
            summary = summary.groupby("circle").agg({
                "red_call_flag": "sum", "enc1_flag": "sum", "enc2_flag": "sum"
            }).reset_index()

            summary = summary.rename(columns={
                "circle": "Circle", "red_call_flag": "Red Call",
                "enc1_flag": "Encroaching1", "enc2_flag": "Encroaching2"
            })

            summary["Platter1"] = summary["Red Call"] + summary["Encroaching1"]
            summary["Platter2"] = summary["Platter1"] + summary["Encroaching2"]
            
            col_order = ["Circle", "Red Call", "Encroaching1", "Platter1","Encroaching2", "Platter2"]
            summary = summary[col_order]
            
            # Droping the rows where Platter2 is zero
            summary= summary[summary["Platter2"] != 0]

            # Sort by Platter1 descending BEFORE adding Total row
            summary = summary.sort_values("Platter1", ascending=False).reset_index(drop=True)

            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)
            
            # Convert DataFrame to list of lists
            data_to_write = [summary.columns.tolist()] + summary.astype(str).values.tolist()
            worksheet.update(data_to_write)

            show_popup("billing code platter report created successfully!", type = "success")
    except Exception as e:
        print(f"Error in billing code status platter function: {e}")
        show_popup(f"Error in billing code status platter function : {e}", type = "error")

def pdna_status_platter(merged_data):
    try:
        spreadsheet = connect_gsheet()

        # To write data in google sheet
        # Open or create the worksheet
        try:
            worksheet = spreadsheet.worksheet("PDNA")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("PDNA", rows=5000, cols=30)
            
        # Clear existing data and write fresh
        worksheet.clear()

        merged_data = merged_data[merged_data["status_code"].str.upper().str.strip() == "PART_DECLARED_NOT_AVAILABLE"]
        
        if not merged_data.empty:
            summary = merged_data.groupby("circle").agg({
                "red_call_flag": "sum", "enc1_flag": "sum", "enc2_flag": "sum"
            }).reset_index()

            summary = summary.rename(columns={
                "circle": "Circle", "red_call_flag": "Red Call",
                "enc1_flag": "Encroaching1", "enc2_flag": "Encroaching2"
            })

            summary["Platter1"] = summary["Red Call"] + summary["Encroaching1"]
            summary["Platter2"] = summary["Platter1"] + summary["Encroaching2"]
            
            col_order = ["Circle", "Red Call", "Encroaching1", "Platter1","Encroaching2", "Platter2"]
            summary = summary[col_order]

            # Droping the rows where Platter2 is zero
            summary= summary[summary["Platter2"] != 0]

            # Sort by Platter1 descending BEFORE adding Total row
            summary = summary.sort_values("Platter1", ascending=False).reset_index(drop=True)
            
            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)
            
            # Convert DataFrame to list of lists
            data_to_write = [summary.columns.tolist()] + summary.astype(str).values.tolist()
            worksheet.update(data_to_write)
            
            show_popup("PDNA platter report created successfully!", type = "success")
    except Exception as e:
        print(f"Error in pdna_status_platter function is : {e}")
        show_popup(f"Error in pdna_status_platter function : {e}", type = "error")
    
def ran_cn_due_status_platter(merged_data):
    try:
        spreadsheet = connect_gsheet()

        # To write data in google sheet
        # Open or create the worksheet
        try:
            worksheet = spreadsheet.worksheet("RAN_CN_DUE")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("RAN_CN_DUE", rows=5000, cols=30)

        # Clear existing data and write fresh
        worksheet.clear()
        
        merged_data = merged_data[(merged_data["status_code"].str.upper().str.strip() == "RAN_C_CN_DUE") | (merged_data["status_code"].str.upper().str.strip() == "RAN_D_CN_DUE")]
        
        if not merged_data.empty:
            summary = merged_data.groupby("circle").agg({
                "red_call_flag": "sum", "enc1_flag": "sum", "enc2_flag": "sum"
            }).reset_index()

            summary = summary.rename(columns={
                "circle": "Circle", "red_call_flag": "Red Call",
                "enc1_flag": "Encroaching1", "enc2_flag": "Encroaching2"
            })

            summary["Platter1"] = summary["Red Call"] + summary["Encroaching1"]
            summary["Platter2"] = summary["Platter1"] + summary["Encroaching2"]
            
            col_order = ["Circle", "Red Call", "Encroaching1", "Platter1","Encroaching2", "Platter2"]
            summary = summary[col_order]

            # Droping the rows where Platter2 is zero
            summary= summary[summary["Platter2"] != 0]

            # Sort by Platter1 descending BEFORE adding Total row
            summary = summary.sort_values("Platter1", ascending=False).reset_index(drop=True)
            
            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)

            
            # Convert DataFrame to list of lists
            data_to_write = [summary.columns.tolist()] + summary.astype(str).values.tolist()
            worksheet.update(data_to_write)
            
            show_popup("RAN_CN_DUE report created successfully!", type = "success")
    except Exception as e:
        print(f"Error in ran_cn_due_status_platter function : {e}")
        show_popup(f"Error in RAN_CN_DUE report function : {e}", type = "error")

def dealerwise_platter(merged_data):
    try:
        spreadsheet = connect_gsheet()

        # To write data in google sheet
        # Open or create the worksheet
        try:
            worksheet = spreadsheet.worksheet("Dealer Platter")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet("Dealer Platter", rows=5000, cols=30)

        # Clear existing data and write fresh
        worksheet.clear()

        filtered_data = merged_data[
        (merged_data["status_code"].str.upper().str.strip().isin(["RAN_C_CN_DUE", "RAN_D_CN_DUE"])) |
        (~merged_data["status_code"].str.upper().str.strip().isin(["RAN_C_CN_DUE", "RAN_D_CN_DUE"]) & 
        (merged_data["customer_type"].str.lower().str.strip() == "dealer"))
        ]

        if not filtered_data.empty:
            summary = filtered_data.groupby("circle").agg({
                "red_call_flag": "sum", "enc1_flag": "sum", "enc2_flag": "sum"
            }).reset_index()

            summary = summary.rename(columns={
                "circle": "Circle", "red_call_flag": "Red Call",
                "enc1_flag": "Encroaching1", "enc2_flag": "Encroaching2"
            })

            summary["Platter1"] = summary["Red Call"] + summary["Encroaching1"]
            summary["Platter2"] = summary["Platter1"] + summary["Encroaching2"]
            
            col_order = ["Circle", "Red Call", "Encroaching1", "Platter1","Encroaching2", "Platter2"]
            summary = summary[col_order]
            
            # Droping the rows where Platter2 is zero
            summary= summary[summary["Platter2"] != 0]

            # Sort by Platter1 descending BEFORE adding Total row
            summary = summary.sort_values("Platter1", ascending=False).reset_index(drop=True)


            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)
            
            # Convert DataFrame to list of lists
            data_to_write = [summary.columns.tolist()] + summary.astype(str).values.tolist()
            worksheet.update(data_to_write)

            show_popup("dealerwise_platter report created successfully!", type = "success")
    except Exception as e:
        print(f"Error in dealerwise_platter function : {e}")
        show_popup(f"Error in dealerwise_platter report function : {e}", type = "error")

def apply_formatting(workbook, worksheet, summary, title_text):
    try:
        fmt_header     = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        fmt_main_title = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'font_size': 14})
        fmt_red        = workbook.add_format({'bold': True, 'bg_color': "#FF0000", 'border': 1, 'align': 'center'})
        fmt_orange     = workbook.add_format({'bg_color': "#FFBF00", 'border': 1, 'align': 'center', 'bold': True})
        fmt_green      = workbook.add_format({'bg_color': "#8BF58B", 'border': 1, 'align': 'center', 'bold': True})
        fmt_white      = workbook.add_format({'border': 1, 'align': 'center', 'bold': True})

        # Single format for entire Total row
        fmt_total = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': '#FFFFFF', 'border': 1, 'align': 'center'})

        # 1. Write Main Title
        report_date = datetime.now().strftime("%d-%b-%Y")
        worksheet.merge_range('A1:F1', f"{title_text} --- {report_date}", fmt_main_title)

        total_row_idx = len(summary) - 1

        # 2. Apply Data Formatting
        for row_num in range(2, len(summary) + 2):
            df_row_idx = row_num - 2
            is_total   = (df_row_idx == total_row_idx)

            for col_num in range(6):
                value = summary.iloc[df_row_idx, col_num]

                if is_total:
                    # Same color for all 6 columns in total row
                    worksheet.write(row_num, col_num, value, fmt_total)
                else:
                    fmt = [fmt_white, fmt_red, fmt_orange, fmt_green, fmt_orange, fmt_green][col_num]
                    worksheet.write(row_num, col_num, value, fmt)

        # 3. Write Column Headers
        for col_num, value in enumerate(summary.columns.values):
            worksheet.write(1, col_num, value, fmt_header)

        worksheet.set_column('A:F', 18)

    except Exception as e:
        print(f"Error in Formatting {worksheet} is : {e}")
        show_popup(f"Error in Formatting {worksheet} is : {e}", type="error")

def apply_tracker_excel_formatting(workbook, worksheet, df, title_text):
    try:
        # --- Formats ---
        fmt_title   = workbook.add_format({'bold': True, 'bg_color': "#EA98B9", 'border': 1, 'align': 'center', 'font_size': 14})
        fmt_header  = workbook.add_format({'bold': True, 'bg_color': "#8CB2ED", 'border': 1, 'align': 'center'})
        fmt_default = workbook.add_format({'bold': True, 'bg_color': "#F2E7DE", 'border': 1, 'align': 'center'})
        fmt_green   = workbook.add_format({'bg_color': '#8BF58B', 'border': 1, 'align': 'center', 'bold': True})
        fmt_blue    = workbook.add_format({'bg_color': '#63B3ED', 'border': 1, 'align': 'center', 'bold': True})

        # Single format for entire Total row
        fmt_total   = workbook.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': '#FFFFFF', 'border': 1, 'align': 'center'})

        headers  = df.columns.tolist()
        num_cols = len(headers)
        total_row_idx = len(df) - 1   # Last row = Total row

        # --- Merge title row across all columns ---
        last_col_letter = chr(ord('A') + num_cols - 1)
        report_date = datetime.now().strftime("%d-%b-%Y")
        worksheet.merge_range(f'A1:{last_col_letter}1', f"{title_text} --- {report_date}", fmt_title)

        # --- Write headers in row 2 ---
        for col_idx, col_name in enumerate(headers):
            worksheet.write(1, col_idx, col_name, fmt_header)

        # --- Find P1 and P2 column indices ---
        p1_cols = [i for i, h in enumerate(headers) if h.endswith("P1")]
        p2_cols = [i for i, h in enumerate(headers) if h.endswith("P2")]

        # --- Write data rows with conditional formatting ---
        for row_idx in range(len(df)):
            is_total = (row_idx == total_row_idx)   # Check if Total row

            for col_idx in range(num_cols):
                value = df.iloc[row_idx, col_idx]

                # Total row → same color for all columns
                if is_total:
                    fmt = fmt_total

                elif col_idx in p1_cols:
                    pair_idx = p1_cols.index(col_idx)
                    p2_idx   = p2_cols[pair_idx] if pair_idx < len(p2_cols) else None
                    try:
                        p1_val = int(float(value))
                        p2_val = int(float(df.iloc[row_idx, p2_idx])) if p2_idx else 1
                    except:
                        p1_val, p2_val = 1, 1

                    if p1_val == 0 and p2_val == 0:
                        fmt = fmt_green
                    elif p1_val == 0:
                        fmt = fmt_blue
                    else:
                        fmt = fmt_default

                elif col_idx in p2_cols:
                    pair_idx = p2_cols.index(col_idx)
                    p1_idx   = p1_cols[pair_idx] if pair_idx < len(p1_cols) else None
                    try:
                        p2_val = int(float(value))
                        p1_val = int(float(df.iloc[row_idx, p1_idx])) if p1_idx else 1
                    except:
                        p1_val, p2_val = 1, 1

                    if p1_val == 0 and p2_val == 0:
                        fmt = fmt_green
                    # elif p1_val == 0:
                    #     fmt = fmt_blue
                    else:
                        fmt = fmt_default

                else:
                    fmt = fmt_default

                worksheet.write(row_idx + 2, col_idx, str(value), fmt)

        # --- Set column widths ---
        worksheet.set_column(0, 0, 12)
        worksheet.set_column(1, 1, 15)
        worksheet.set_column(2, num_cols - 1, 12)

    except Exception as e:
        print(f"Error in apply_tracker_excel_formatting: {e}")
        show_popup(f"Error in tracker Excel formatting: {e}", type="error")

def fetch_and_format_report():
    try:
        spreadsheet = connect_gsheet()
        output = io.BytesIO()

        # Regular sheets with standard formatting
        sheet_configs = [
            {"sheet": "Daily Circlewise Platter", "title": "All Circlewise Platter And Targets"},
            {"sheet": "Statuswise Platter",        "title": "Statuswise Platter And Targets"},
            {"sheet": "Billing Code",              "title": "Billing Code Problem"},
            {"sheet": "PDNA",                      "title": "PDNA"},
            {"sheet": "RAN_CN_DUE",                "title": "RAN_C/D_CN_DUE Calls On The Platter And Targets"},
            {"sheet": "Dealer Platter",            "title": "Dealer Circlewise Platter And Targets"},
        ]

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:

            # --- Standard sheets ---
            for config in sheet_configs:
                sheet_name = config["sheet"]
                title_text = config["title"]
                try:
                    ws       = spreadsheet.worksheet(sheet_name)
                    raw_data = ws.get_all_values()

                    if not raw_data or len(raw_data) < 2:
                        print(f"No data found in sheet: {sheet_name}")
                        continue

                    df = pd.DataFrame(raw_data[1:], columns=raw_data[0])

                    # Convert numeric columns
                    for col in df.columns[1:]:
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

                    df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
                    apply_formatting(writer.book, writer.sheets[sheet_name], df, title_text)
                    print(f"✅ Sheet written: {sheet_name}")

                except gspread.WorksheetNotFound:
                    print(f"⚠️ Sheet not found, skipping: {sheet_name}")
                    continue

            # --- Tracker sheet (special formatting) ---
            try:
                tracker_ws   = spreadsheet.worksheet("Tracker")
                tracker_data = tracker_ws.get_all_values()

                if tracker_data and len(tracker_data) >= 2:
                    tracker_df = pd.DataFrame(tracker_data[1:], columns=tracker_data[0])

                    tracker_df.to_excel(writer, sheet_name="Tracker", index=False, startrow=1)

                    # ✅ Use special tracker formatting
                    apply_tracker_excel_formatting(
                        writer.book,
                        writer.sheets["Tracker"],
                        tracker_df,
                        "Daily Tracker"
                    )
                    print("✅ Tracker sheet written")
                else:
                    print("⚠️ No data in Tracker sheet")

            except gspread.WorksheetNotFound:
                print("⚠️ Tracker sheet not found, skipping")

        show_popup("Platter report ready to download!", type="success")
        return output.getvalue()

    except Exception as e:
        print(f"Error in fetch_and_format_report: {e}")
        show_popup(f"Error generating report: {e}", type="error")
        return None

# # --- Main Execution ---
# if __name__ == "__main__":
#     if len(sys.argv) < 2:
#         print("Please provide the raw file path as an argument.")
#     else:
#         raw_file = sys.argv[1]
        # status_file = r"D:\Amstrad\Service\Statuswise_norms_teams_data.xlsx"
        # out_name = f"service_daily_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        # final_df = func1(raw_file)
        # print("The shape of final_df is :", final_df.shape)
        # if final_df is not None:
        #     # We open the writer ONCE here and pass it to all functions
        #     with pd.ExcelWriter(out_name, engine="xlsxwriter") as writer:
        #         circlewise_platter(final_df, writer)
        #         statuswise_platter(final_df, writer)
        #         billing_code_status_platter(final_df, writer)
        #         pdna_status_platter(final_df,writer)
        #         ran_cn_due_status_platter(final_df,writer)
        #         dealerwise_status_platter(final_df,writer)
        #     print(f"Process Complete. Report saved: {out_name}")
