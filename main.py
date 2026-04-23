import sys
import pandas as pd
from datetime import datetime

def func1(raw_file, statuswise_file):
    try:
        data = pd.read_excel(raw_file)
        data.columns = data.columns.str.lower().str.replace(" ","_").str.replace(".", "_").str.strip()
        
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
        
        status_data = pd.read_excel(statuswise_file)
        status_data.columns = status_data.columns.str.lower().str.strip().str.replace(" ", "_")

        merged_data = data.merge(status_data[["status","team", "number"]], left_on= "status_code", right_on="status", how= "left")

        # Adding filter on teams
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
        merged_data.to_excel("new_one.xlsx", index = False)
        return merged_data 

    except Exception as e:
        print(f"Error in func1: {e}")
        return e

def apply_formatting(workbook, worksheet, summary, title_text):
    try:
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        fmt_main_title = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'font_size': 14})
        fmt_red = workbook.add_format({'bold': True, 'bg_color': "#FF0000", 'border': 1, 'align': 'center'})
        fmt_orange = workbook.add_format({'bg_color': "#FFBF00", 'border': 1, 'align': 'center', 'bold': True})
        fmt_green = workbook.add_format({'bg_color': "#8BF58B", 'border': 1, 'align': 'center', 'bold': True})
        fmt_white = workbook.add_format({'border': 1, 'align': 'center', 'bold': True})

        # 1. Write Main Title
        report_date = datetime.now().strftime("%d-%b-%Y")
        worksheet.merge_range('A1:F1', f"{title_text} --- {report_date}", fmt_main_title)

        # 2. Apply Data Formatting
        for row_num in range(2, len(summary) + 2):
            worksheet.write(row_num, 0, summary.iloc[row_num-2, 0], fmt_white) # Label Column
            worksheet.write(row_num, 1, summary.iloc[row_num-2, 1], fmt_red)   # Red Call
            worksheet.write(row_num, 2, summary.iloc[row_num-2, 2], fmt_orange)# Enc1
            worksheet.write(row_num, 3, summary.iloc[row_num-2, 3], fmt_green) # Platter1
            worksheet.write(row_num, 4, summary.iloc[row_num-2, 4], fmt_orange)# Enc2
            worksheet.write(row_num, 5, summary.iloc[row_num-2, 5], fmt_green) # Platter2

        # 3. Write Column Headers
        for col_num, value in enumerate(summary.columns.values):
            worksheet.write(1, col_num, value, fmt_header)

        worksheet.set_column('A:F', 18)
    except Exception as e:
        print(f"Error in Formatting {worksheet} is : {e}")
        return f"Error in formatting {worksheet} is : {e}"

def circlewise_platter(merged_data, writer):
    try:

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
        
        totals = summary.select_dtypes(include='number').sum()
        total_row = pd.DataFrame([totals])
        total_row["Circle"] = "Total"
        summary = pd.concat([summary, total_row], ignore_index=True)

        summary.to_excel(writer, sheet_name="Daily Circlewise Platter", index=False, startrow=1)
        apply_formatting(writer.book, writer.sheets["Daily Circlewise Platter"], summary, "All Circlewise Platter And Targets")
    except Exception as e:
        print(f"Error in circlewise platter function: {e}")
        return f"Error in circlewise platter function: {e}"

def statuswise_platter(merged_data, writer):

    try:
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
        
        totals = summary.select_dtypes(include='number').sum()
        total_row = pd.DataFrame([totals])
        total_row["Status"] = "Total"
        summary = pd.concat([summary, total_row], ignore_index=True)

        summary.to_excel(writer, sheet_name="Statuswise Platter", index=False, startrow=1)
        apply_formatting(writer.book, writer.sheets["Statuswise Platter"], summary, "Statuswise Platter And Targets")
    except Exception as e:
        print(f"Error in statuswise_platter function : {e}")
        return f"Error in statuswise_platter function : {e}"

def billing_code_status_platter(merged_data, writer):
    try:
        summary = merged_data[merged_data["status_code"].str.upper().str.strip() == "BILLING_CODE_PROBLEM"]
        
        # if merged_data.empty:
            # worksheet = writer.book.add_worksheet("Billing Code")
            # fmt_msg = writer.book.add_format({'bold': True, 'font_color': 'red', 'font_size': 12})
            # worksheet.write('A1', "No data found for Billing Code Problem.", fmt_msg)
            # return
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
            
            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)

            summary.to_excel(writer, sheet_name="Billing Code", index=False, startrow=1)
            apply_formatting(writer.book, writer.sheets["Billing Code"], summary, "Billing Code Problem")
    except Exception as e:
        print(f"Error in billing code status platter function: {e}")
        return f"Error in billing code status platter function: {e}"

def pdna_status_platter(merged_data, writer):
    try:
        merged_data = merged_data[merged_data["status_code"].str.upper().str.strip() == "PDNA"]
        
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
            
            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)

            summary.to_excel(writer, sheet_name="PDNA", index=False, startrow=1)
            apply_formatting(writer.book, writer.sheets["PDNA"], summary, "PDNA")
    except Exception as e:
        print(f"Error in pdna_status_platter function is : {e}")
        return f"Error in pdna_status_platter function is : {e}"
    
def ran_cn_due_status_platter(merged_data, writer):
    try:
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
            
            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)

            summary.to_excel(writer, sheet_name="RAN_CN_DUE", index=False, startrow=1)
            apply_formatting(writer.book, writer.sheets["RAN_CN_DUE"], summary, "RAN_C/D_CN_DUE Calls On The Platter And Targets")
    except Exception as e:
        print(f"Error in ran_cn_due_status_platter function : {e}")
        return f"Error in ran_cn_due_status_platter function : {e}"

def dealerwise_status_platter(merged_data,writer):
    try:
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
            
            totals = summary.select_dtypes(include='number').sum()
            total_row = pd.DataFrame([totals])
            total_row["Circle"] = "Total"
            summary = pd.concat([summary, total_row], ignore_index=True)

            summary.to_excel(writer, sheet_name="Dealer Platter", index=False, startrow=1)
            apply_formatting(writer.book, writer.sheets["Dealer Platter"], summary, "Dealer Circlewise Platter And Targets")
    except Exception as e:
        print(f"Error in dealerwise_status_platter function : {e}")
        return f"Error in dealerwise_status_platter function : {e}"

# --- Main Execution ---
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Please provide the raw file path as an argument.")
    else:
        raw_file = sys.argv[1]
        # status_file = r"D:\Amstrad\Service\Statuswise_norms_teams_data.xlsx"
        # out_name = f"service_daily_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        # final_df = func1(raw_file, status_file)

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