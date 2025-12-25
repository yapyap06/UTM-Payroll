import pandas as pd
import openpyxl
import warnings
from datetime import datetime
import os

warnings.simplefilter(action='ignore', category=FutureWarning)

def load_data(employee_file):
    df=None

    if os.path.exists(employee_file):
        try:
            df=pd.read_excel(employee_file)
            print(f"\n{'=' * 50}")
            print(f"{'EMPLOYEE DATABASE':^50}")  # center align
            print(f"{'=' * 50}")
            print(df.to_string(index=False))
            print(f"{'=' * 50}\n")

        except Exception as e:
            print(f"[!] Error reading file: {e}")
    else:
        print(f"\n[!] File '{employee_file}' not found. Please add an employee first.\n")

    return df


def write_to_salary_file(salary_file,employee_id,normal_pay,ot_pay,total_gross_pay,epf,socso,after_deduct_pay):
    # can create file if file doesn't exist
    if os.path.exists(salary_file):
        df = pd.read_excel(salary_file)
    else:
        df = pd.DataFrame(
            columns=['employee_id', 'datetime', 'normal_hours_paid', 'ot_hours_paid', 'gross_pay', 'epf', 'socso',
                     'net_paid'])

    new_row = pd.DataFrame({'employee_id': [employee_id],
                            'datetime': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],  # excel time format (string)
                            'normal_hours_paid': [normal_pay],
                            'ot_hours_paid': [ot_pay],
                            'gross_pay': [total_gross_pay],
                            'epf':[epf],
                            'socso':[socso],
                            'net_paid':[after_deduct_pay]})

    df = pd.concat([df, new_row], ignore_index=True)

    try:
        df.to_excel(salary_file, index=False)
        print(f"\n[+] New Salary Record for Employee ID: {employee_id} has been saved.\n")
    except PermissionError:
        print(
            f"\n[!] Error: Could not save '{salary_file}'. Is the file open? Please close it.\n")  # if excel file opened at background


def working_hours(employee_file):
    df=pd.read_excel(employee_file)
    id= str(input("Enter Employee ID: ")).strip()

    index_row_update = df[df['employee_id'].astype(str) == id]

    if len(index_row_update) == 0:
        print(f"\n[!] Employee {id} was not found!\n")  # fstring
        return None, None, None #no crash

    row_index = index_row_update.index[0]
    standard_hours = df.loc[row_index, "standard_hours"]

    # time - input Error handling
    while True:
        try:
            actual_work_time = float(input(f"Input actual work time for ID {id}: "))
            break
        except ValueError:
            print("[!] Invalid input. Please enter a number.")


    if actual_work_time>standard_hours:
        normal_hours=standard_hours
        ot_hours=actual_work_time-standard_hours
    else:
        normal_hours=actual_work_time
        ot_hours=0

    print(f"\n{'WORK SUMMARY':^30}")  # fstr
    print(f"{'-' * 30}")
    print(f"Normal Hours : {normal_hours:.1f} hrs")
    print(f"Overtime     : {ot_hours:.1f} hrs")
    print(f"{'-' * 30}\n")

    return normal_hours, ot_hours, id

def gross_pay(employee_file,normal_hours,ot_hours,employee_id):
    df=pd.read_excel(employee_file)

    row_index = df[df['employee_id'].astype(str) == employee_id].index[0]

    normal_pay=normal_hours*(df.loc[row_index,"hourly_rate"])
    ot_pay=ot_hours*(df.loc[row_index,'overtime_rate'])
    total_gross_pay=normal_pay+ot_pay

    print(f"{'PAYMENT BREAKDOWN':^30}")
    print(f"{'-' * 30}")
    print(f"Normal Pay   : RM {normal_pay:>8.2f}")
    print(f"OT Pay       : RM {ot_pay:>8.2f}")
    print(f"{'-' * 30}")
    print(f"Gross Pay    : RM {total_gross_pay:>8.2f}")
    print(f"{'=' * 30}\n")

    return normal_pay,ot_pay,total_gross_pay

def deductions(total_gross_pay):
    EPF = 0.11
    SOCSO = 0.005

    EPF_deduct=total_gross_pay*EPF
    SOCSO_deduct=total_gross_pay*SOCSO

    print(f"{'DEDUCTIONS':^30}")
    print(f"{'-' * 30}")
    print(f"EPF (11%)    : RM {EPF_deduct:>8.2f}")
    print(f"SOCSO (0.5%) : RM {SOCSO_deduct:>8.2f}")

    after_deduct_pay=total_gross_pay-EPF_deduct-SOCSO_deduct
    return EPF_deduct,SOCSO_deduct,after_deduct_pay


def net_salary(after_deduct_pay):
    print(f"{'-' * 30}")
    print(f"NET SALARY   : RM {after_deduct_pay:>8.2f}")  #
    print(f"{'=' * 30}\n")

def generate_payslip(employee_file,salary_file):
    if not os.path.exists(employee_file):
        print("\n[!] No employee database found.\n")
        return

    df=pd.read_excel(employee_file)
    print(f"\n{'=' * 40}")  # fstr
    print(f"{'GENERATE PAYSLIP':^40}")
    print(f"{'=' * 40}")

    result = working_hours(employee_file)

    if result[0] is None:
        return

    normal_hours, ot_hours, employee_id = result

    row_index = df[df['employee_id'].astype(str) == employee_id].index[0]  # astype double check str
    name = df.loc[row_index, 'name']
    print(f"Generating Slip for: {name} (ID: {employee_id})")

    normal_pay,ot_pay,total_gross_pay = gross_pay(employee_file,normal_hours,ot_hours,employee_id)
    epf,socso,after_deduct_pay= deductions(total_gross_pay)

    print("After Deduct EPF and SOCSO:")
    net_salary(after_deduct_pay)
    write_to_salary_file(salary_file, employee_id, normal_pay, ot_pay, total_gross_pay, epf, socso, after_deduct_pay)


def add_employee(employee_file):
    if os.path.exists(employee_file):
        df = pd.read_excel(employee_file)
    else:
        df = pd.DataFrame(columns=['employee_id', 'name', 'hourly_rate', 'standard_hours', 'overtime_rate'])

    print(f"\n{'ADD NEW EMPLOYEE':^30}")
    print(f"{'-' * 30}")

    id=str(input("Employee ID: ")).strip()

    if not df.empty and len(df[df['employee_id'].astype(str) == id]) > 0:
        print(f"\n[!] Employee ID {id} is already in the system!\n")
        return

    name=input("Employee Name: ").strip()

    while True:
        try:
            rate = float(input("Hourly Rate (RM): "))
            break
        except ValueError:
            print("[!] Invalid input! Please enter a number.")

    while True:
        try:
            std_hour = float(input("Standard Working Hours: "))
            break
        except ValueError:
            print("[!] Invalid input! Please enter a number.")

    while True:
        try:
            ot_rate = float(input("Overtime Rate (RM): "))
            break
        except ValueError:
            print("[!] Invalid input! Please enter a number.")

    new_row=pd.DataFrame({'employee_id':[id],
                          'name':[name],
                          'hourly_rate':[rate],
                          'standard_hours':[std_hour],
                          'overtime_rate':[ot_rate]})

    df=pd.concat([df,new_row], ignore_index=True)

    try:
        df.to_excel(employee_file, index=False)
        print(f"\n[+] {name} has been added to the system successfully!\n")  # fstr
    except PermissionError:
        print(f"\n[!] Error: Could not save file. Please close '{employee_file}'.\n")


def remove_employee(employee_file):
    if not os.path.exists(employee_file):
        print("\n[!] Database not found.\n")
        return

    df=pd.read_excel(employee_file)

    id=str(input("Employee you want to remove (ID): ")).strip()

    index_to_remove=df[df['employee_id'].astype(str)==id].index

    if len(index_to_remove)==0:
        print(f"\n[!] Employee {id} was not found!\n")
        return

    print(f"\n[?] Warning: This action is irreversible.")
    confirm_remove = input(f"Type 'Y' to confirm removing Employee {id}: ").upper()

    if confirm_remove=='Y':
        df=df.drop(index_to_remove)
        try:
            df.to_excel(employee_file, index=False)
            print(f"\n[+] Employee {id} has successfully been removed.\n")  # fstr
        except PermissionError:
            print(f"\n[!] Error: Could not save file. Close '{employee_file}'.\n")
    else:
        print("\n[i] Action canceled.\n")


def update_employee(employee_file):
    if not os.path.exists(employee_file):
        print("\n[!] Database not found.\n")
        return

    df=pd.read_excel(employee_file)

    id_row=str(input("Which employee's info you want to update(ID): ")).strip()
    index_row_update = df[df['employee_id'].astype(str)==id_row]

    if len(index_row_update) == 0:
        print(f"\n[!] Employee {id_row} was not found!\n")
        return

    row_index=index_row_update.index[0]
    valid_cols=['name','hourly_rate','standard_hours','overtime_rate']

    print(f"\nAvailable columns: {valid_cols}")
    column_name = input("Column to update: ").lower()

    if column_name not in valid_cols:
        print(f"\n[!] Column '{column_name}' does not exist!\n")
        return

    info=input(f"Type your update info for {column_name}: ")

    if column_name in ['hourly_rate', 'standard_hours', 'overtime_rate']:
        try:
            info = float(info)
        except ValueError:
            print(f"\n[!] Error: {column_name} requires a number. Update cancelled.\n")
            return

    df.at[row_index,column_name]=info

    try:
        df.to_excel(employee_file, index=False)
        print(f"\n[+] Info '{column_name}' for Employee {id} updated to {info}.\n")  # fstr
    except PermissionError:
        print(f"\n[!] Error: Could not save file. Close '{employee_file}'.\n")


def main():
    employee_file="employee_data.xlsx"
    salary_file="payslip_data.xlsx"
    try:
        while True:
            print(f"{'=' * 40}")  # fstr
            print(f"{'PAYROLL MANAGEMENT SYSTEM':^40}")
            print(f"{'=' * 40}")
            print(" [A] Load File")
            print(" [B] Add Employee")
            print(" [C] Remove Employee")
            print(" [D] Update Employee Info")
            print(" [E] Generate Pay Slip")
            print(" [X] Exit")
            print("-" * 40)

            choice=input("Select an option: ").strip().upper()

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
                print("\nExiting System. Goodbye.🫡💪\n")
                break
            else:
                print("\n[!] Invalid choice. Please try again.\n")

    except KeyboardInterrupt as err:
        print("\n\nExit Successfully")


if __name__=="__main__":
    main()