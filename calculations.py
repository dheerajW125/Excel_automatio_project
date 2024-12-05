import pandas as pd
from pymongo import MongoClient

# Function 1: Extract row-wise data from Excel
def get_row_wise_data_with_headers(file_path, sheet_name="BILLING"):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    required_columns = ["SL NO", "Employee Name","DEPT","Code","Designation", "TOTAL PAID DAYS", "Canteen Deduction", "OT HRS"]
    
    if not all(col in df.columns for col in required_columns):
        missing_columns = [col for col in required_columns if col not in df.columns]
        raise ValueError(f"Missing columns: {missing_columns}")
    
    df["OT HRS"] = pd.to_numeric(df["OT HRS"], errors='coerce').fillna(0).astype(float)
    df["TOTAL PAID DAYS"] = pd.to_numeric(df["TOTAL PAID DAYS"], errors='coerce').fillna(0).astype(float)
    
    row_wise_data = []
    for _, row in df.iterrows():
        if pd.isna(row["SL NO"]):  # Stop processing if SL NO is empty
            break
        row_data = {
            "SL NO": row["SL NO"],
            "DEPT": row["DEPT"],
            "Code": row["Code"],
            "Employee Name": row["Employee Name"],
            "Designation": row["Designation"],
            "TOTAL PAID DAYS": row["TOTAL PAID DAYS"],
            "Canteen Deduction": row["Canteen Deduction"],
            "OT HRS": row["OT HRS"]
        }
        row_wise_data.append(row_data)
    # print(row_wise_data)
    return row_wise_data

# Function 2: Calculate wages
client = MongoClient("mongodb://localhost:27017/")
db = client["Tann_DB"]
collection = db["data"]

def base_salary_month_cal(designation, no_of_days, ot_hrs, canteen_days):
    price_fetch = collection.find_one({'Monthly Wages (30 /31 days)': {'$exists': True}})
    wage_rate = {}
    wage_earned = {}
    unskilled = 1259
    semi_skilled = 1355
    skilled_iti = 1462
    dme_2_yrs = 1578
    supervisor =1578
    machine_oprtr =1461
    maintain = 1000
    quality_insp = 1000

    try:
        wage = price_fetch['Monthly Wages (30 /31 days)'][0][designation]
        if designation == "Unskilled":
            bonus = unskilled
        elif designation == "Semi Skilled":
            bonus = semi_skilled
        elif designation == "Skilled (ITI)":
            bonus = skilled_iti
        elif designation == "DME (2 Yrs)":
            bonus = dme_2_yrs
        elif designation == "Supervisor":
            bonus = supervisor
        elif designation == "MACHINE OPERATOR":
            bonus = machine_oprtr
        elif designation == "QUALITY INSPECTOR":
            bonus = quality_insp
        elif designation == "MAINTENANCE":
            bonus = maintain

        wage_rate["base salary"] = wage
        wage_rate["St Bonus"] = round(bonus)
        wage_rate["OT Amount"] = ot_hrs
        wage_rate["Special allowance"] = 0
        wage_rate["Leave Wages"] = wage / 26 * 1.25
        
        lev_wages = wage_rate["Leave Wages"] / 30 * no_of_days
        pf = round((min(wage, 15000) * 13 / 100), 1)

        wage_earned["base salary"] = (wage / 30 * no_of_days)
        wage_earned["St Bonus"] = round(bonus / 30 * no_of_days, 1)
        wage_earned["OT Amount"] = round(wage / 26 / 8 * 2 * ot_hrs, 1)
        wage_earned["Special allowance"] = 0
        wage_earned["Leave Wages"] = round(lev_wages, 1)
        wage_earned["pf@13"] = pf
        wage_earned["Uniform"] = 40.80

        wage_earned["Gross Amount"] = (
            wage_earned["base salary"]
            + wage_earned["St Bonus"]
            + wage_earned["OT Amount"]
            + round(lev_wages)
            + pf
        )
        wage_earned["Service charges"] = wage_earned["Gross Amount"] * 9 / 100
        ESI = (
            wage_earned["Gross Amount"] * 3.25 / 100
            if wage - wage_rate["Leave Wages"] - wage_rate["St Bonus"] < 21000
            else 0
        )
        wage_earned["ESI@3.25%"] = round(ESI, 1)
        wage_earned["Sub Total"] = (
            wage_earned["Gross Amount"]
            + round(pf, 1)
            + round(ESI, 1)
            + wage_earned["Uniform"]
        )
        wage_earned["Total"] = wage_earned["Sub Total"] + wage_earned["Service charges"]
        wage_earned["canteen"] = round(canteen_days * 25, 1)
        wage_earned["Billing Amount"] = round(
            wage_earned["Total"] - wage_earned["canteen"], 1
        )
    except Exception as e:
        print(f"Error processing data for designation '{designation}': {e}")
        wage_rate, wage_earned = {}, {}

    return wage_rate, wage_earned

file_path = r"C:\Users\Dheeraj.W\Desktop\MONTH OF APR-24.xlsx"

designation_mapping = {
    "SKILLED": "Skilled (ITI)",
    "DME": "DME (2 Yrs)",
    "ITI OPERATOR": "Skilled (ITI)",
    "ITI": "Skilled (ITI)",
    "UNSKILLED": "Unskilled",
    "Supervisor":"Supervisor",
    "MACHINE OPERATOR":"MACHINE OPERATOR",
    "QUALITY INSPECTOR":"QUALITY INSPECTOR",
    "MAINTENANCE":"MAINTENANCE",
}


try:
    extracted_data = get_row_wise_data_with_headers(file_path)
    for row in extracted_data:
        original_designation = row["Designation"]
        mapped_designation = designation_mapping.get(original_designation.upper(), None)
        
        if mapped_designation is None:
            print(f"Skipping row with invalid designation: {original_designation}")
            continue 

        no_of_days = row["TOTAL PAID DAYS"]
        canteen_days = row["Canteen Deduction"]
        ot_hrs = row["OT HRS"]

        wage_rate, wage_earned = base_salary_month_cal(mapped_designation, no_of_days, ot_hrs, canteen_days)

        print(f"Input Row: {row}")
        print(f"Mapped Designation: {mapped_designation}")
        print(f"Wage Rate: {wage_rate}")
        print(f"Wage Earned: {wage_earned}")
        print("-" * 50)

except ValueError as e:
    print(f"Error: {e}")
except FileNotFoundError:
    print("Error: The specified file was not found.")
