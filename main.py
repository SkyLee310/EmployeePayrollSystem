import pandas as pd
import openpyxl

#initialize variable
EPF=0.11
SOCSO=0.005

def load_data(filename):
    df=None
    try:
        df=pd.read_excel(filename)
    except FileNotFoundError:
        print("File does not found! Check whether filepath is valid.")

    if df is not None:
        print("Employee Database")
        print(df.head())

    return df

def working_hours():
    pass

def gross_pay():
    pass

def deductions():
    pass

def net_salary():
    pass

def generate_payslip():
    pass

def add_employee():
    pass

def remove_employee():
    pass

def update_employee():
    pass

def employee_info():
    pass

def main():
    load_data("employee_data.xlsx")

if __name__=="__main__":
    main()