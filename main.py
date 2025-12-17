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

def working_hours(filename):
    df=pd.read_excel(filename)
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

def gross_pay():
    pass

def deductions():
    pass

def net_salary():
    pass

def generate_payslip(filename):

    print("---------Generate Payslip System-----------")
    normal_hours, ot_hours, employee_id = working_hours(filename)
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
        df.to_excel(filename, index=False)
    else:
        print("The action has been canceled\n")


def update_employee(filename):

    df=pd.read_excel(filename)

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
    df.to_excel(filename, index=False)


def main():
    filename="employee_data.xlsx"
    while True:
        print("Employee Payroll Management System")
        print("A.Load File\nB.Adding Employee\nC.Remove Employee\nD.Update Employee's Information\nE.Generate Pay Slip\n"
              "To Exit, Type X")
        choice=input("What action you want to do: ")
        if choice=='A':
            load_data(filename)
        elif choice=='B':
            add_employee(filename)
        elif choice=='C':
            remove_employee(filename)
        elif choice=='D':
            update_employee(filename)
        elif choice=='E':
            pass #generate_payslip(filename)
        elif choice=='X':
            break


if __name__=="__main__":
    main()