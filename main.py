from fileinput import filename

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
        print("\nEmployee Database\n")
        print(df.head(),"\n")


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
    print(f"{name} has been added to the system!\n\n")

    df.to_excel(filename, index=False)


def remove_employee(filename):

    df=pd.read_excel(filename)

    id=input("Employee you want to remove (ID): ")
    index_to_remove=df[df['employee_id'].astype(str)==id].index

    if len(index_to_remove)==0:
        print(f"Employee {id} was not found!\n")
        return

    print(f"Warning this action is irreversible. Do you sure want to remove employee {id} ?")
    confirm_remove=input("Type Y to confirm, Type anything else to cancel: ")
    if confirm_remove=='Y':
        df=df.drop(index_to_remove)
        print(f"Employee {id} has successfully been removed\n")
        df.to_excel(filename, index=False)
    else:
        print("The action has been canceled\n")



def update_employee():
    pass

def employee_info():
    pass

def main():
    filename="employee_data.xlsx"
    while True:
        print("Employee Payroll Management System")
        print("A.Load File\nB.Adding Employee\nC.Remove Employee\nTo Exit, Type X")
        choice=input("What action you want to do: ")
        if choice=='A':
            load_data(filename)
        elif choice=='B':
            add_employee(filename)
        elif choice=='C':
            remove_employee(filename)
        elif choice=='X':
            break


if __name__=="__main__":
    main()