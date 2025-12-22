import pandas as pd
import openpyxl
import warnings
from datetime import datetime
warnings.simplefilter(action='ignore', category=FutureWarning)

def load_data(employee_file):
    df=None
    try:
        df=pd.read_excel(employee_file)
    except FileNotFoundError:
        print("File does not found! Check whether filepath is valid.")

    if df is not None:
        print("\n------------------Employee Database-------------------")
        print(df.head(),"\n")


    return df

def write_to_salary_file(salary_file,employee_id,normal_pay,ot_pay,total_gross_pay,epf,socso,after_deduct_pay):
    df=pd.read_excel(salary_file)

    new_row = pd.DataFrame({'employee_id': [employee_id],
                            'datetime': [datetime.now()],
                            'normal_hours_paid': [normal_pay],
                            'ot_hours_paid': [ot_pay],
                            'gross_pay': [total_gross_pay],
                            'epf':[epf],
                            'socso':[socso],
                            'net_paid':[after_deduct_pay]})

    df = pd.concat([df, new_row], ignore_index=True)
    print(f"New Salary Record for Employee ID:{employee_id} has been added to the system!\n\n")

    df.to_excel(salary_file, index=False)


def working_hours(employee_file):
    df=pd.read_excel(employee_file)
    id=input("Which employee working hours need to be calculate(ID): ")

    index_row_update = df[df['employee_id'].astype(str) == id]

    if len(index_row_update) == 0:
        print(f"Employee {id} was not found!\n")
        return

    row_index = index_row_update.index[0]

    standard_hours = df.loc[row_index, "standard_hours"]
    actual_work_time=float(input("Input actual work time: "))


    if actual_work_time>standard_hours:
        normal_hours=standard_hours
        ot_hours=actual_work_time-standard_hours
    else:
        normal_hours=actual_work_time
        ot_hours=0

    print(f"Normal Pay Work Time: {normal_hours:.1f} hours")
    print(f"Overtime Pay Work Time: {ot_hours:.1f} hours\n\n")

    return normal_hours, ot_hours, id

def gross_pay(employee_file,normal_hours,ot_hours,employee_id):
    df=pd.read_excel(employee_file)

    row_index = df[df['employee_id'].astype(str) == employee_id].index[0]

    normal_pay=normal_hours*(df.loc[row_index,"hourly_rate"])
    ot_pay=ot_hours*(df.loc[row_index,'overtime_rate'])
    total_gross_pay=normal_pay+ot_pay

    print(f"Normal Hours Paid: RM{normal_pay:.2f}")
    print(f"OT Hours Paid: RM{ot_pay:.2f}")
    print(f"Total Gross Pay: RM{total_gross_pay:.2f}")

    return normal_pay,ot_pay,total_gross_pay

def deductions(total_gross_pay):
    EPF = 0.11
    SOCSO = 0.005

    EPF_deduct=total_gross_pay*EPF
    SOCSO_deduct=total_gross_pay*SOCSO
    print(f"Deduct EPF RM{EPF_deduct:.2f}")
    print(f"Deduct SOCSO RM{SOCSO_deduct:.2f}")
    after_deduct_pay=total_gross_pay-EPF_deduct-SOCSO_deduct
    return EPF_deduct,SOCSO_deduct,after_deduct_pay


def net_salary(after_deduct_pay):
    print(f"Your Net Salary is RM{after_deduct_pay:.2f}\n\n")

def generate_payslip(employee_file,salary_file):
    df=pd.read_excel(employee_file)
    print("---------Generate Payslip System-----------")
    normal_hours, ot_hours, employee_id = working_hours(employee_file)
    row_index = df[df['employee_id'].astype(str) == employee_id].index[0]
    print(f"Employee's {df.loc[row_index,'name']} Payslip")
    normal_pay,ot_pay,total_gross_pay = gross_pay(employee_file,normal_hours,ot_hours,employee_id)
    epf,socso,after_deduct_pay= deductions(total_gross_pay)
    print("After Deduct EPF and SOCSO:")
    net_salary(after_deduct_pay)
    write_to_salary_file(salary_file, employee_id, normal_pay, ot_pay, total_gross_pay, epf, socso, after_deduct_pay)


def add_employee(employee_file):

    df=pd.read_excel(employee_file)

    print("Add New Employee")
    id=input("Employee ID: ")
    index_to_add = df[df['employee_id'].astype(str) == id].index

    if len(index_to_add) != 0:
        print(f"Employee {id} was already there!\n")
        return

    name=input("Employee Name: ").strip()
    while True:
        try:
            rate = float(input("Hourly Rate: "))
            break
        except ValueError:
            print("Invalid input! Please enter a number for Hourly Rate.")

    while True:
        try:
            std_hour = float(input("Standard Working Hours: "))
            break
        except ValueError:
            print("Invalid input! Please enter a number for Working Hours.")

    while True:
        try:
            ot_rate = float(input("Overtime Rate: "))
            break
        except ValueError:
            print("Invalid input! Please enter a number for Overtime Rate.")

    new_row=pd.DataFrame({'employee_id':[id],
                          'name':[name],
                          'hourly_rate':[rate],
                          'standard_hours':[std_hour],
                          'overtime_rate':[ot_rate]})

    df=pd.concat([df,new_row], ignore_index=True)
    print(f"{name} has been added to the system!\n\n")

    df.to_excel(employee_file, index=False)


def remove_employee(employee_file):

    df=pd.read_excel(employee_file)

    id=int(input("Employee you want to remove (ID): "))
    index_to_remove=df[df['employee_id']==id].index

    if len(index_to_remove)==0:
        print(f"Employee {id} was not found!\n")
        return

    print(f"Warning this action is irreversible. Do you sure want to remove employee {id} ?")
    confirm_remove=input("Type Y to confirm, Type anything else to cancel: ")
    if confirm_remove=='Y':
        df=df.drop(index_to_remove)
        print(f"Employee {id} has successfully been removed\n")
        df.to_excel(employee_file, index=False)
    else:
        print("The action has been canceled\n")


def update_employee(employee_file):

    df=pd.read_excel(employee_file)

    id_row=int(input("Which employee's info you want to update(ID): "))
    index_row_update = df[df['employee_id']==id_row]

    if len(index_row_update) == 0:
        print(f"Employee {id_row} was not found!\n")
        return

    row_index=index_row_update.index[0]

    valid_cols=['name','hourly_rate','standard_hours','overtime_rate']
    print(f"Select a column to update:{valid_cols}")
    column_name=input("Column: ")

    if column_name not in valid_cols:
        print(f"{column_name} are not exist!\n")
        return

    info=input(f"Type your update info for {column_name}: ")

    if column_name in ['hourly_rate', 'standard_hours', 'overtime_rate']:
        try:
            info = float(info)
        except ValueError:
            print(f"Error: {column_name} requires a number. Update cancelled.")
            return

    df.loc[row_index,column_name]=info

    print(f"Info {column_name} for Employee {id_row} has updated to {info}\n\n")
    df.to_excel(employee_file, index=False)


def main():
    employee_file="employee_data.xlsx"
    salary_file="payslip_data.xlsx"
    try:
        while True:
            print("Employee Payroll Management System")
            print("A.Load File\nB.Adding Employee\nC.Remove Employee\nD.Update Employee's Information\nE.Generate Pay Slip\n"
                  "To Exit, Type X")
            choice=input("What action you want to do: ").upper()
            if choice=='A':
                load_data(employee_file)
            elif choice=='B':
                add_employee(employee_file)
            elif choice=='C':
                remove_employee(employee_file)
            elif choice=='D':
                update_employee(employee_file)
            elif choice=='E':
                generate_payslip(employee_file,salary_file)
            elif choice=='X':
                break
    except KeyboardInterrupt as err:
        print("\nExit Successfully")


if __name__=="__main__":
    main()