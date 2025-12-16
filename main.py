import pandas as pd
import openpyxl
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

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

def add_employee(filename):

    df=pd.read_excel(filename)

    print("Add New Employee")
    id=input("Employee ID: ")
    name=input("Employee Name: ")
    rate=float(input("Hourly Rate: "))
    std_hour=float(input("Standard Working Hours: "))
    ot_rate=float(input("Overtime Rate: "))

    new_row=pd.DataFrame({'employee_id':[id],
                          'name':[name],
                          'hourly_rate':[rate],
                          'standard_hours':[std_hour],
                          'overtime_rate':[ot_rate]})

    df=pd.concat([df,new_row], ignore_index=True)
    print(f"{name} has been added to the system!")

    df.to_excel(filename, index=False)


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