from random import randint
import os
from openpyxl import Workbook
from tkinter import *
from tkinter import ttk, messagebox
from datetime import *
import sqlite3
from tkcalendar import Calendar
import calendar


class DatabaseConnection:
    def __init__(self, db):
        self.conn = sqlite3.connect(db)
        self.cur = self.conn.cursor()

        # Create employee table
        self.conn.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            monthly_salary INT,
            overtime_rate INT,
            allowances INT)
        ''')

        # Create salary table
        self.conn.execute('''
        CREATE TABLE IF NOT EXISTS salary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER,
            absent INT,
            holidays INT,
            overtime_hours INT,
            no_pay INT,
            base_pay INT,
            gross_pay INT,
            salaried_month DATE,
            FOREIGN KEY (employee_id) REFERENCES employees (id))
        ''')

        self.conn.commit()

    def __del__(self):
        self.conn.close()


class PayrollDB(DatabaseConnection):

    def fetch(self):
        self.cur.execute("SELECT * FROM employees")
        rows = self.cur.fetchall()
        return rows

    def insert(self, name, monthly_salary, overtime_rate, allowances):
        self.cur.execute("INSERT INTO employees VALUES (NULL, ?, ?, ?, ?)",
                         (name, monthly_salary, overtime_rate, allowances))
        self.conn.commit()

    def update(self, name, monthly_salary, overtime_rate, allowances, employee_id):
        self.cur.execute(
            "UPDATE employees SET name = ?, monthly_salary = ?, overtime_rate = ?, allowances = ? WHERE id = ?",
            (name, monthly_salary, overtime_rate, allowances, employee_id))
        self.conn.commit()

    def delete(self, employee_id):
        self.cur.execute("DELETE FROM employees WHERE id = ?", (employee_id,))
        self.conn.commit()

    def search(self, employee_id):
        self.cur.execute("SELECT * FROM employees WHERE id = ?",
                         (employee_id,))
        row = self.cur.fetchone()
        return row

    def record_payroll(self, employee_id, absent, holiday, overtime_hours, no_pay, base_pay, gross_pay, salaried_month):
        self.conn.execute('''
            INSERT INTO salary (employee_id, absent, holidays, overtime_hours, no_pay, base_pay, 
            gross_pay, salaried_month)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (employee_id, absent, holiday, overtime_hours, no_pay, base_pay, gross_pay, salaried_month))
        self.conn.commit()

    def get_monthly_salary_report(self, employee_id):
        self.cur.execute('''
        SELECT salaried_month, base_pay, no_pay, gross_pay
        FROM salary
        WHERE employee_id = ?
        ORDER BY salaried_month DESC, employee_id
        ''', (employee_id,))
        rows = self.cur.fetchall()

        # Converting the rows to a list of dictionaries
        report = []
        for row in rows:
            report.append({
                'salaried_month': row[0],
                'base_pay': row[1],
                'no_pay': row[2],
                'gross_pay': row[3]
            })

        return report

    def get_overall_salary_summary(self, employee_id, start_month, end_month):
        self.cur.execute('''
        SELECT salaried_month, SUM(base_pay), SUM(no_pay), SUM(gross_pay)
        FROM salary
        WHERE employee_id = ? AND salaried_month >= ? AND salaried_month <= ?
        GROUP BY salaried_month
        ORDER BY salaried_month DESC, employee_id
        ''', (employee_id, start_month, end_month))
        rows = self.cur.fetchall()

        # Converting the rows to a list of dictionaries
        report = []
        for row in rows:
            report.append({
                'salaried_month': row[0],
                'base_pay': row[1],
                'no_pay': row[2],
                'gross_pay': row[3]
            })

        return report

    def get_salary_values_for_date_range(self, start_month, end_month):
        self.cur.execute('''
        SELECT employee_id, salaried_month, no_pay, base_pay, gross_pay
        FROM salary
        WHERE salaried_month >= ? AND salaried_month <= ?
        ORDER BY salaried_month DESC, employee_id
        ''', (start_month, end_month))
        rows = self.cur.fetchall()

        # Converting the rows to a list of dictionaries
        report = []
        for row in rows:
            report.append({
                'employee_id': row[0],
                'salaried_month': row[1],
                'no_pay': row[2],
                'base_pay': row[3],
                'gross_pay': row[4]
            })

        return report

    def get_column_names(self,):
        self.cur.execute("PRAGMA table_info(salary)")
        columns = self.cur.fetchall()
        return columns

    def generate_salary_entries(self, entry_gen_start_date, entry_gen_end_date):
        current_date = entry_gen_start_date

        while current_date < entry_gen_end_date:
            for i in range(50):
                # Generate random data for each salary entry
                employee_id = randint(1, 16)
                absent = randint(0, 3)
                holiday = randint(0, 5)
                overtime_hours = randint(0, 20)
                base_pay = randint(1500, 10000)
                no_pay = randint(0, 750)
                gross_pay = base_pay - no_pay

                # Set the salaried_month to the 28th of the current month
                salaried_month = current_date.replace(day=28).strftime('%Y-%m-%d')

                # Check if an entry already exists for the employee_id and salaried_month
                self.cur.execute('''
                    SELECT employee_id, salaried_month FROM salary WHERE employee_id = ? AND salaried_month = ?
                ''', (employee_id, salaried_month))

                existing_entry = self.cur.fetchone()

                if existing_entry:
                    print(f"Entry already exists for employee_id: {employee_id} and salaried_month: {salaried_month}")
                    continue

                self.conn.execute('''
                    INSERT INTO salary (employee_id, absent, holidays, overtime_hours, no_pay,
                    base_pay, gross_pay, salaried_month)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (employee_id, absent, holiday, overtime_hours, no_pay, base_pay, gross_pay, salaried_month))

            # Move to the next month
            current_date = current_date.replace(day=1) + timedelta(days=32)

        self.conn.commit()


# Initializing the database
payroll_db = PayrollDB('./grifindo_payroll.db')

# Check if the employees table is empty
payroll_db.cur.execute("SELECT COUNT(*) FROM employees")
result = payroll_db.cur.fetchone()

if result[0] == 0:
    # Populating the Employee table
    payroll_db.insert('John Doe', 5000, 15, 1000)
    payroll_db.insert('Jane Smith', 4500, 12, 800)
    payroll_db.insert('Michael Johnson', 6000, 20, 1500)
    payroll_db.insert('Sarah Williams', 5500, 18, 1200)
    payroll_db.insert('Robert Brown', 4000, 10, 600)
    payroll_db.insert('Emily Davis', 4800, 14, 900)
    payroll_db.insert('David Anderson', 5200, 16, 1100)
    payroll_db.insert('Jennifer Wilson', 5100, 16, 1050)
    payroll_db.insert('Daniel Thompson', 4700, 12, 850)
    payroll_db.insert('Olivia Garcia', 5200, 15, 1050)
    payroll_db.insert('Matthew Martinez', 6000, 20, 1500)
    payroll_db.insert('Sophia Clark', 4500, 13, 800)
    payroll_db.insert('James Rodriguez', 4900, 14, 950)
    payroll_db.insert('Lily Lewis', 5100, 17, 1050)
    payroll_db.insert('Benjamin Young', 5300, 17, 1150)

# Populating the Salary table

payroll_db.cur.execute("SELECT COUNT(*) FROM salary")
result = payroll_db.cur.fetchone()

if result[0] == 0:
    start_date = datetime(2018, 1, 1)
    end_date = datetime(2023, 6, 1)
    payroll_db.generate_salary_entries(start_date, end_date)


class MonthPicker:
    def __init__(self, parent):
        self.parent = parent
        self.cal = None

    def select_month(self, title, value_var, label_value, button, label_frame):
        def set_month():
            selected_date = self.cal.selection_get()
            value_var.set(selected_date.strftime("%Y-%m-%d"))
            button.grid_remove()
            label_value.grid(padx=10, pady=10)
            top.destroy()

        top = Toplevel(self.parent)
        top.title(title)

        self.cal = Calendar(top, selectmode="day", date_pattern="yyyy-mm-dd")
        self.cal.pack(padx=10, pady=10)

        confirm_button = ttk.Button(top, text="Select Date", command=set_month)
        confirm_button.pack(padx=10, pady=10)


class EmployeeComponent(Frame):
    def __init__(self, parent):

        # Initializing the parent class (Frame) and passing the parent parameter to it
        super().__init__(parent)

        self.employee_frame = LabelFrame(self, text="Employee Details")
        self.employee_frame.pack(padx=10, pady=10)

        self.name_label = Label(self.employee_frame, text="Name: ")
        self.name_label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.name_entry = Entry(self.employee_frame)
        self.name_entry.grid(row=0, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.salary_label = Label(self.employee_frame, text="Monthly Salary: ")
        self.salary_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.salary_entry = Entry(self.employee_frame)
        self.salary_entry.grid(row=1, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.overtime_label = Label(self.employee_frame, text="Overtime Rate (hourly): ")
        self.overtime_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
        self.overtime_entry = Entry(self.employee_frame)
        self.overtime_entry.grid(row=2, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.allowances_label = Label(self.employee_frame, text="Allowances: ")
        self.allowances_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
        self.allowances_entry = Entry(self.employee_frame)
        self.allowances_entry.grid(row=3, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.register_button = ttk.Button(self.employee_frame, text="Register", command=self.register_employee)
        self.register_button.grid(row=4, column=0, padx=10, pady=10, ipadx=10, sticky=W + E)

        self.update_button = ttk.Button(self.employee_frame, text="Update", command=self.update_employee)
        self.update_button.grid(row=4, column=1, padx=10, pady=10, ipadx=10)

        self.clear_button = ttk.Button(self.employee_frame, text="Clear", command=self.clear_entry)
        self.clear_button.grid(row=4, column=2, padx=10, pady=10, ipadx=10, sticky=W)

        self.db_output_frame = LabelFrame(self, text="Registered Employee")
        self.db_output_frame.pack(padx=10, pady=10, fill='both')

        self.db_output_label_info = Label(self.db_output_frame, text='Select an employee to update')
        self.db_output_label_info.pack(padx=10, pady=3, anchor=W)

        self.db_output_label = Label(self.db_output_frame,
                                     text='Employee ID | Name | Salary | Overtime Rate | Allowances')
        self.db_output_label.pack(padx=10, pady=3)

        self.db_output = Listbox(self.db_output_frame)
        self.db_output.pack(side=LEFT, fill='both', expand=True, pady=10)

        self.scrollbar = ttk.Scrollbar(self.db_output_frame, orient='vertical', command=self.db_output.yview)
        self.scrollbar.pack(side=RIGHT, fill='y')

        # Configure the Listbox to use the scrollbar
        self.db_output.config(yscrollcommand=self.scrollbar.set)

        self.db_output.bind('<<ListboxSelect>>', self.select_entry)

        self.populate_list()

        self.delete_frame = LabelFrame(self, text="Delete Employee")
        self.delete_frame.pack(padx=10, pady=10, fill='both')

        self.search_label = Label(self.delete_frame, text="Employee ID: ")
        self.search_label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.search_entry = Entry(self.delete_frame)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5)

        self.delete_button = ttk.Button(self.delete_frame, text="Delete Employee", command=self.delete_employee,
                                        )
        self.delete_button.grid(row=0, column=2, padx=5, pady=5)

    def populate_list(self):
        # All employee records will be fetched from the DB and displayed in the textbox

        self.db_output.delete(0, END)
        # Loop through records
        for row in PayrollDB.fetch(payroll_db):
            # Insert into list
            self.db_output.insert(END, row)

    def register_employee(self):

        if self.name_entry.get() == '' or self.salary_entry.get() == '' or self.overtime_entry.get() == '' or \
                self.allowances_entry.get() == '':
            messagebox.showerror("Required Fields", 'Please fill all the fields, to register the employee!')
            return

        else:
            PayrollDB.insert(payroll_db, self.name_entry.get(), self.salary_entry.get(), self.overtime_entry.get(),
                             self.allowances_entry.get())
            # Clear the entry fields after registration
            self.name_entry.delete(0, END)
            self.salary_entry.delete(0, END)
            self.overtime_entry.delete(0, END)
            self.allowances_entry.delete(0, END)

            messagebox.showinfo('Success', 'Employee registered successfully!')
            self.populate_list()
            return

    def update_employee(self):

        if self.name_entry.get() == '' or self.salary_entry.get() == '' or self.overtime_entry.get() == '' or \
                self.allowances_entry.get() == '':
            messagebox.showerror("Required Fields", 'Please fill all the fields, to update the employee!')
            return

        else:
            PayrollDB.update(payroll_db, self.name_entry.get(), self.salary_entry.get(), self.overtime_entry.get(),
                             self.allowances_entry.get(), self.search_entry.get())
            # Clear the entry fields after registration
            self.name_entry.delete(0, END)
            self.salary_entry.delete(0, END)
            self.overtime_entry.delete(0, END)
            self.allowances_entry.delete(0, END)
            self.search_entry.delete(0, END)

            messagebox.showinfo('Success', 'Employee record updated successfully!')
            self.populate_list()
            return

    def select_entry(self, event):
        # This function also works as a search. As data from the selected entry will populate the entry boxes.

        selection = self.db_output.curselection()
        if selection:
            index = selection[0]
            selected_entry = self.db_output.get(index)

            self.search_entry.delete(0, END)
            self.search_entry.insert(END, selected_entry[0])

            self.name_entry.delete(0, END)
            self.name_entry.insert(END, selected_entry[1])

            self.salary_entry.delete(0, END)
            self.salary_entry.insert(END, selected_entry[2])

            self.overtime_entry.delete(0, END)
            self.overtime_entry.insert(END, selected_entry[3])

            self.allowances_entry.delete(0, END)
            self.allowances_entry.insert(END, selected_entry[4])

    def clear_entry(self):

        self.search_entry.delete(0, END)
        self.name_entry.delete(0, END)
        self.salary_entry.delete(0, END)
        self.overtime_entry.delete(0, END)
        self.allowances_entry.delete(0, END)

    def delete_employee(self):

        if self.search_entry.get() == '':
            messagebox.showerror("Required Fields", 'Please enter a valid ID to delete the employee!')
            return

        else:
            PayrollDB.delete(payroll_db, self.search_entry.get())
            # Clear the entry fields after registration
            self.search_entry.delete(0, END)

            messagebox.showinfo('Success', 'Employee deleted successfully!')
            self.populate_list()
            return


class SettingsComponent(Frame):
    def __init__(self, parent):
        # Initializing the parent class (Frame) and passing the parent parameter to it
        super().__init__(parent)

        self.salary_cycle_date_range = None
        self.cycle_days_entry_label = None
        self.leave_limit_entry_label = None
        self.first_day = StringVar()
        self.end_day = StringVar()
        self.month_picker = MonthPicker(self)

        self.settings_frame = LabelFrame(self, text="Settings for Salary Calculation")
        self.settings_frame.pack(padx=10, pady=10,fill='both')

        self.today = date.today()
        self.first_day = date(self.today.year, self.today.month, 1)
        self.last_day = date(self.today.year, self.today.month,
                             calendar.monthrange(self.today.year, self.today.month)[1])

        self.setting_start_date_label = Label(self.settings_frame, text="Salary Cycle Start Date (current month): ")
        self.setting_start_date_label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.settings_start_date_label_value = Label(self.settings_frame, text=str(self.first_day))
        self.settings_start_date_label_value.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        self.cycle_end_label = Label(self.settings_frame, text="Salary Cycle End Date (current month): ")
        self.cycle_end_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.settings_end_date_label_value = Label(self.settings_frame, text=str(self.last_day))
        self.settings_end_date_label_value.grid(row=1, column=1, padx=5, pady=5, sticky=W)

        self.cycle_days_label = Label(self.settings_frame, text="Salary Cycle Days: ")
        self.cycle_days_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
        self.cycle_days_entry = Entry(self.settings_frame)

        # Calculating the number of days for the current month
        self.salary_cycle_date_range = (self.last_day - self.first_day).days + 1

        self.cycle_days_entry.insert(0, str(self.salary_cycle_date_range))
        self.cycle_days_entry.grid(row=2, column=1, padx=5, pady=5)

        self.leave_limit_label = Label(self.settings_frame, text="Leave Limit (per year): ")
        self.leave_limit_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
        self.leave_limit_entry = Entry(self.settings_frame)
        self.leave_limit_entry.insert(0, '30')
        self.leave_limit_entry.grid(row=3, column=1, padx=5, pady=5)

        self.update_button = ttk.Button(self.settings_frame, text="Confirm Salary Settings",
                                        command=self.update_settings,
                                        )
        self.update_button.grid(row=4, column=0, padx=5, pady=5, sticky=W + E)

        self.reset_button = ttk.Button(self.settings_frame, text="Reset", command=self.reset_settings,
                                       )
        self.reset_button.grid(row=4, column=1, padx=5, pady=5, sticky=W + E)

    def reset_settings(self):
        # Clear the entry fields
        self.cycle_days_entry.delete(0, END)
        self.leave_limit_entry.delete(0, END)
        self.cycle_days_entry.insert(0, str(self.salary_cycle_date_range))  # Setting the default value
        self.leave_limit_entry.insert(0, '2')  # Setting the default value

        # Reset the grid layout
        self.cycle_days_entry.grid(row=2, column=1, padx=5, pady=5)
        self.leave_limit_entry.grid(row=3, column=1, padx=5, pady=5)

        # Remove the labels if they exist
        if self.cycle_days_entry_label:
            self.cycle_days_entry_label.grid_remove()
            self.cycle_days_entry_label = None

        if self.leave_limit_entry_label:
            self.leave_limit_entry_label.grid_remove()
            self.leave_limit_entry_label = None

    def update_settings(self):
        cycle_days = self.cycle_days_entry.get()
        leave_limit = self.leave_limit_entry.get()

        # Remove the entry fields
        self.cycle_days_entry.grid_remove()
        self.leave_limit_entry.grid_remove()

        # Create and display labels with the entered values
        self.cycle_days_entry_label = Label(self.settings_frame, text=cycle_days)
        self.cycle_days_entry_label.grid(row=2, column=1, padx=5, pady=5)

        self.leave_limit_entry_label = Label(self.settings_frame, text=leave_limit)
        self.leave_limit_entry_label.grid(row=3, column=1, padx=5, pady=5)

        messagebox.showinfo("Success", "Settings updated successfully!")


class SalaryComponent(SettingsComponent):
    def __init__(self, parent):
        super().__init__(parent)

        self.employee_id = StringVar()
        self.employee_name = StringVar()
        self.employee_salary = StringVar()
        self.employee_overtime = StringVar()
        self.employee_allowances = StringVar()
        self.month_picker = MonthPicker(self)
        self.no_pay_value = 0
        self.base_pay_value = 0
        self.gross_pay = 0

        self.employee_frame = LabelFrame(self, text="Employee Details")
        self.employee_frame.pack(padx=10, pady=10,fill='both')

        self.search_label = Label(self.employee_frame, text="Employee ID: ")
        self.search_label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.search_entry = Entry(self.employee_frame)
        self.search_entry.grid(row=0, column=1, padx=5, pady=5, sticky=W + E)

        self.search_button = ttk.Button(self.employee_frame, text="Search", command=self.search_employee,
                                        )
        self.search_button.grid(row=0, column=2, padx=10, pady=10)

        self.id_label = Label(self.employee_frame, text="Employee ID: ")
        self.id_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.employee_id_label = Label(self.employee_frame, textvariable=self.employee_id,
                                       )
        self.employee_id_label.grid(row=1, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.name_label = Label(self.employee_frame, text="Employee Name: ")
        self.name_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
        self.employee_name_label = Label(self.employee_frame, textvariable=self.employee_name,
                                         )
        self.employee_name_label.grid(row=2, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.salary_label = Label(self.employee_frame, text="Monthly Salary: ")
        self.salary_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
        self.employee_salary_label = Label(self.employee_frame, textvariable=self.employee_salary,
                                           )
        self.employee_salary_label.grid(row=3, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.overtime_label = Label(self.employee_frame, text="Overtime Rate (hourly): ")
        self.overtime_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
        self.employee_overtime_label = Label(self.employee_frame, textvariable=self.employee_overtime,
                                             )
        self.employee_overtime_label.grid(row=4, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.allowances_label = Label(self.employee_frame, text="Allowances: ")
        self.allowances_label.grid(row=5, column=0, padx=5, pady=5, sticky=W)
        self.employee_allowances_label = Label(self.employee_frame, textvariable=self.employee_allowances,
                                               )
        self.employee_allowances_label.grid(row=5, column=1, columnspan=3, padx=10, pady=5, sticky=W + E)

        self.salary_frame = LabelFrame(self, text="Calculate Salary")
        self.salary_frame.pack(padx=10, pady=10)

        self.start_date_label = Label(self.salary_frame, text="Start Date: ")
        self.start_date_label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.start_date_value = StringVar()
        self.start_date_label_value = Label(self.salary_frame, textvariable=self.start_date_value)
        self.start_date_label_value.grid(row=0, column=1, padx=5, pady=5, sticky=W)

        self.start_date_button = ttk.Button(self.salary_frame, text="Select Start Date", command=self.select_start_date)
        self.start_date_button.grid(row=0, column=1, padx=5, pady=5)

        self.end_date_label = Label(self.salary_frame, text="End Date: ")
        self.end_date_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.end_date_value = StringVar()
        self.end_date_label_value = Label(self.salary_frame, textvariable=self.end_date_value)
        self.end_date_label_value.grid(row=1, column=1, padx=5, pady=5, sticky=W)

        self.end_date_button = ttk.Button(self.salary_frame, text="Select End Date", command=self.select_end_date)
        self.end_date_button.grid(row=1, column=1, padx=5, pady=5)

        self.absent_days_label = Label(self.salary_frame, text="No. of Absent Days: ")
        self.absent_days_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
        self.absent_days_entry = Entry(self.salary_frame)
        self.absent_days_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=5)

        self.holidays_label = Label(self.salary_frame, text="No. of Holidays: ")
        self.holidays_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
        self.holidays_entry = Entry(self.salary_frame)
        self.holidays_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=5)

        self.overtime_hours_label = Label(self.salary_frame, text="No. of Overtime Hours: ")
        self.overtime_hours_label.grid(row=5, column=0, padx=5, pady=5, sticky=W)
        self.overtime_hours_entry = Entry(self.salary_frame)
        self.overtime_hours_entry.grid(row=5, column=1, columnspan=2, padx=5, pady=5)

        self.reset_button = ttk.Button(self.salary_frame, text="Reset All Entries", command=self.reset_layout)
        self.reset_button.grid(row=6, column=0, padx=5, pady=5, sticky=W + E)

        self.calculate_button = ttk.Button(self.salary_frame, text="Calculate Salary", command=self.calculate_salary)
        self.calculate_button.grid(row=6, column=1, columnspan=2, padx=5, pady=5, sticky=W + E)

        self.calculate_button = ttk.Button(self.salary_frame, text="Record Employee Payroll",command=self.record_payroll)
        self.calculate_button.grid(row=7, columnspan=2, padx=10, pady=5, sticky=W + E)

    def select_start_date(self):
        self.month_picker.select_month(
            "Select Start Date",
            self.start_date_value,
            self.start_date_label_value,
            self.start_date_button,
            self.salary_frame
        )

    def select_end_date(self):

        self.month_picker.select_month(
            "Select End Date",
            self.end_date_value,
            self.end_date_label_value,
            self.end_date_button,
            self.salary_frame
        )

        start_date_str = self.start_date_value.get()
        end_date_str = self.end_date_value.get()

        if start_date_str and end_date_str:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()

            if start_date > end_date:
                messagebox.showerror("Date Error!", "End date must be greater than start date")

    def search_employee(self):
        employee_id = self.search_entry.get()

        if employee_id:
            output = PayrollDB.search(payroll_db, employee_id)

            if output:
                self.employee_id.set(output[0])
                self.employee_name.set(output[1])
                self.employee_salary.set(output[2])
                self.employee_overtime.set(output[3])
                self.employee_allowances.set(output[4])
            else:
                messagebox.showerror("Employee Not Found", "Employee not found!")
        else:
            messagebox.showerror("Required Fields", "Please enter an valid Employee ID!")

        self.search_entry.delete(0, END)

    def calculate_salary(self):
        if self.absent_days_entry.get() == "" or self.holidays_entry.get() == "" or \
                self.overtime_hours_entry.get() == "" or self.start_date_value.get() == "" or \
                self.end_date_value.get() == "":
            messagebox.showerror("Required Fields", "Please fill in all fields")
        elif self.employee_id.get() == "":
            messagebox.showerror("Required Fields", "Search for an employee to calculate salary!")
        else:
            # Get input values
            no_of_absent_days = int(self.absent_days_entry.get())
            no_of_overtime_hours = int(self.overtime_hours_entry.get())

            # Perform salary calculation
            monthly_salary = int(self.employee_salary.get())
            allowances = int(self.employee_allowances.get())
            overtime_rate = int(self.employee_overtime.get())

            # Calculate salary cycle date range
            if self.cycle_days_entry.get():
                salary_cycle_date_range = int(self.cycle_days_entry.get())
            else:
                salary_cycle_date_range = int(self.salary_cycle_date_range)

            # Calculate no-pay value
            self.no_pay_value = (monthly_salary / salary_cycle_date_range) * no_of_absent_days

            # Calculate base pay value
            self.base_pay_value = monthly_salary + allowances + (overtime_rate * no_of_overtime_hours)

            # Calculate gross pay
            government_tax_rate = 0.25  # Example value, you may need to adjust it
            self.gross_pay = self.base_pay_value - (self.no_pay_value + self.base_pay_value * government_tax_rate)

            # Display the calculated values
            messagebox.showinfo("Salary Calculation", f"Base Pay: {self.base_pay_value}\nGross Pay: {self.gross_pay}")

    def record_payroll(self):
        if self.absent_days_entry.get() == "" or self.holidays_entry.get() == "" or \
                self.overtime_hours_entry.get() == "":
            messagebox.showerror("Required Fields", "Please fill in all fields")
        elif self.employee_id.get() == "":
            messagebox.showerror("Required Fields", "Search for an employee to calculate salary!")
        else:
            salaried_month = date.today()

            # Access the class-level variables
            no_pay = self.no_pay_value
            base_pay = self.base_pay_value
            gross_pay = self.gross_pay

            PayrollDB.record_payroll(payroll_db, self.employee_id.get(), self.absent_days_entry.get(),
                                     self.holidays_entry.get(), self.overtime_hours_entry.get(), no_pay, base_pay,
                                     gross_pay, salaried_month)

            messagebox.showinfo('Success', 'Employee Payroll Recorded Successfully!')

    def reset_layout(self):
        self.search_entry.delete(0, END)
        self.employee_id.set("")
        self.employee_name.set("")
        self.employee_salary.set("")
        self.employee_overtime.set("")
        self.employee_allowances.set("")
        self.start_date_value.set("")
        self.end_date_value.set("")
        self.start_date_label_value.grid_remove()
        self.start_date_button.grid(row=0, column=1, padx=5, pady=5)
        self.end_date_label_value.grid_remove()
        self.end_date_button.grid(row=1, column=1, padx=5, pady=5)


class ReportGenerator(Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.month_picker = MonthPicker(self)
        self.start_date_value = StringVar()
        self.end_date_value = StringVar()

        self.employee_report_frame = Frame(self)
        self.employee_report_frame.pack()

        self.report_type_frame = LabelFrame(self.employee_report_frame, text="Report Type")
        self.report_type_frame.pack(padx=10, pady=10,fill='both')

        self.report_type_var = IntVar()

        self.radio_button_monthly_salary = ttk.Radiobutton(self.report_type_frame, text="Monthly Salary Report",
                                                           variable=self.report_type_var, value=1)
        self.radio_button_monthly_salary.grid(row=0, column=0, padx=10, pady=10, sticky=W)
        self.radio_button_overall_summary = ttk.Radiobutton(self.report_type_frame, text="Overall Salary Summary",
                                                            variable=self.report_type_var, value=2)
        self.radio_button_overall_summary.grid(row=0, column=1, padx=10, pady=10, sticky=W)
        self.radio_button_gross_pay = ttk.Radiobutton(self.report_type_frame, text="Gross Pay Report",
                                                      variable=self.report_type_var, value=3)
        self.radio_button_gross_pay.grid(row=1, columnspan=2, padx=10, pady=10)

        self.employee_report_label_frame = LabelFrame(self.employee_report_frame, text="Employee Report")
        self.employee_report_label_frame.pack(padx=10, pady=10,fill='both')

        self.employee_id_label = Label(self.employee_report_label_frame, text="Employee ID: ")
        self.employee_id_label.grid(row=0, column=0, padx=10, pady=10, sticky=W)
        self.employee_id_entry = Entry(self.employee_report_label_frame)
        self.employee_id_entry.grid(row=0, column=1, padx=10, pady=10)

        self.report_start_date_label = Label(self.employee_report_label_frame,
                                             text="Generate report from (YYYY-MM-DD): ")
        self.report_start_date_label.grid(row=1, column=0, padx=10, pady=10, sticky=W)
        self.report_start_date_button = ttk.Button(self.employee_report_label_frame, text="Select Start Date",
                                                   command=self.report_select_start_date)
        self.report_start_date_button.grid(row=1, column=1, padx=10, pady=10)
        self.report_start_date_label_value = Label(self.employee_report_label_frame, textvariable=self.start_date_value)
        self.report_start_date_label_value.grid(row=1, column=1, padx=5, pady=5, sticky=W)

        self.report_end_date_label = Label(self.employee_report_label_frame, text="Generate report to (YYYY-MM-DD): ")
        self.report_end_date_label.grid(row=2, column=0, padx=10, pady=10, sticky=W)
        self.report_end_date_button = ttk.Button(self.employee_report_label_frame, text="Select Start Date",
                                                 command=self.report_select_end_date)
        self.report_end_date_button.grid(row=2, column=1, padx=10, pady=10)
        self.report_end_date_label_value = Label(self.employee_report_label_frame, textvariable=self.end_date_value)
        self.report_end_date_label_value.grid(row=2, column=1, padx=5, pady=5, sticky=W)


        self.generate_report = ttk.Button(self.employee_report_label_frame, text="Generate Report",
                                          command=self.generate_report)
        self.generate_report.grid(row=3, columnspan=2, padx=10, pady=10,)

        self.reset_button = ttk.Button(self.employee_report_label_frame, text="Reset", command=self.reset_layout)
        self.reset_button.grid(row=3, column=1, padx=10, pady=10, sticky=E)

    def reset_layout(self):
        self.employee_id_entry.delete(0, END)
        self.report_start_date_label_value.grid_remove()
        self.report_end_date_label_value.grid_remove()
        self.report_start_date_button.grid(row=1, column=1, padx=10, pady=10)
        self.report_end_date_button.grid(row=2, column=1, padx=10, pady=10)
        self.report_type_var.set(-1)  # Reset the radio buttons
        self.radio_button_monthly_salary.grid(row=0, column=0, padx=10, pady=10, sticky=W)
        self.radio_button_overall_summary.grid(row=0, column=1, padx=10, pady=10, sticky=W)
        self.radio_button_gross_pay.grid(row=1, columnspan=2, padx=10, pady=10)

    def report_select_start_date(self):
        self.month_picker.select_month(
            "Select Start Date",
            self.start_date_value,
            self.report_start_date_label_value,
            self.report_start_date_button,
            self.employee_report_label_frame
        )

    def report_select_end_date(self):

        self.month_picker.select_month(
            "Select End Date",
            self.end_date_value,
            self.report_end_date_label_value,
            self.report_end_date_button,
            self.employee_report_label_frame
        )

    def validate_fields(self):
        employee_id = self.employee_id_entry.get()
        start_date = self.start_date_value.get()
        end_date = self.end_date_value.get()
        report_type = self.report_type_var.get()

        if report_type == 1 and not employee_id:
            messagebox.showerror("Missing Information",
                                 "Please provide the employee ID to generate the Monthly Salary Report.")
            self.reset_layout()
            return False

        elif report_type == 2 and (not employee_id or not start_date or not end_date):
            messagebox.showerror("Missing Information",
                                 "Please provide all necessary information to generate the Overall Salary Report.")
            return False

        elif report_type == 3 and (not start_date or not end_date):
            messagebox.showerror("Missing Information",
                                 "Please provide date range to generate the Gross Pay Report.")
            return False

        elif report_type == 0 and (not employee_id or not start_date or not end_date):
            messagebox.showerror("Missing Information",
                                 "Please provide all necessary information to generate the report.")
            return False

        return True

    def generate_report(self):
        if not self.validate_fields():
            return

        employee_id = self.employee_id_entry.get()
        start_date = self.start_date_value.get()
        end_date = self.end_date_value.get()
        report_type = self.report_type_var.get()

        if report_type == 1:  # Monthly Salary Report
            report = payroll_db.get_monthly_salary_report(employee_id)
            if report:
                filename = f"emp_id_{employee_id}_monthly_salary_report.xlsx"
                self.export_report(report, filename)
            else:
                messagebox.showinfo("No Data", "No monthly salary report data found.")

        elif report_type == 2:  # Overall Salary Summary
            report = payroll_db.get_overall_salary_summary(employee_id, start_date, end_date)
            if report:
                filename = f"emp_id_{employee_id}_overall_salary_summary_{start_date}_to_{end_date}.xlsx"
                self.export_report(report, filename)
            else:
                messagebox.showinfo("No Data", "No overall salary summary data found.")

        elif report_type == 3:  # Gross Pay Report
            report = payroll_db.get_salary_values_for_date_range(start_date, end_date)
            if report:
                filename = f"gross_pay_report_{start_date}_to_{end_date}.xlsx"
                self.export_report(report, filename)
            else:
                messagebox.showinfo("No Data", "No gross pay report data found.")

    def export_report(self, report, filename):
        workbook = Workbook()
        sheet = workbook.active

        column_names = list(report[0].keys())
        sheet.append(column_names)

        # Write the report data to the workbook
        for row in report:
            sheet.append(list(row.values()))

        # Create the "Report" folder if it doesn't exist
        folder_path = "Reports"
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Save the workbook to the "Report" folder with the specified filename
        file_path = os.path.join(folder_path, filename)
        workbook.save(file_path)

        messagebox.showinfo("Success", f"Report Generated and in the 'Reports' folder\n\nFile Name: {filename}")
        self.reset_layout()

    def get_column_names(self, report):
        if report:
            column_names = list(report[0].keys())
            return column_names
        else:
            return []


class PayrollSystem(Tk):
    def __init__(self):
        super().__init__()
        self.title("Grifindo Toys Payroll System")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(padx=10, pady=10)

        self.employee_component = EmployeeComponent(self.notebook)
        self.salary_component = SalaryComponent(self.notebook)
        self.report_generator = ReportGenerator(self.notebook)

        self.notebook.add(self.employee_component, text="Employee")
        self.notebook.add(self.salary_component, text="Salary")
        self.notebook.add(self.report_generator, text="Report")


if __name__ == "__main__":
    payroll_system = PayrollSystem()
    payroll_system.mainloop()
