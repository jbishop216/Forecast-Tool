import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
from models import Employee, GA01Week, WorkCode, Forecast, get_session
from forecast_calculator import update_forecast
from import_export import import_employees, import_ga01_weeks, export_forecast

class EmployeeDialog(tk.Toplevel):
    def __init__(self, parent, employee=None):
        super().__init__(parent)
        self.employee = employee
        self.result = None
        
        self.title("Employee Details")
        self.geometry("400x300")
        self.resizable(False, False)
        
        # Create form fields
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.name_var, width=30).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Team:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.team_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.team_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Manager Code:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.manager_code_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.manager_code_var, width=30).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Cost Center:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.cost_center_var, width=30).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Start Date (YYYY-MM-DD):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.start_date_var = tk.StringVar()
        self.start_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(frame, textvariable=self.start_date_var, width=30).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="End Date (YYYY-MM-DD):").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.end_date_var = tk.StringVar()
        self.has_end_date_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Has End Date", variable=self.has_end_date_var, 
                        command=self.toggle_end_date).grid(row=5, column=1, sticky=tk.W, pady=5)
        
        self.end_date_entry = ttk.Entry(frame, textvariable=self.end_date_var, width=30, state="disabled")
        self.end_date_entry.grid(row=6, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Employment Type:").grid(row=7, column=0, sticky=tk.W, pady=5)
        self.emp_type_var = tk.StringVar(value="FTE")
        ttk.Combobox(frame, textvariable=self.emp_type_var, values=["FTE", "Contractor"], 
                     width=28, state="readonly").grid(row=7, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Populate fields if editing existing employee
        if employee:
            self.name_var.set(employee.name)
            self.team_var.set(employee.team or "")
            self.manager_code_var.set(employee.manager_code)
            self.cost_center_var.set(employee.cost_center)
            
            if employee.start_date:
                self.start_date_var.set(employee.start_date.strftime("%Y-%m-%d"))
            
            if employee.end_date:
                self.has_end_date_var.set(True)
                self.end_date_var.set(employee.end_date.strftime("%Y-%m-%d"))
                self.end_date_entry.config(state="normal")
            
            self.emp_type_var.set(employee.employment_type)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def toggle_end_date(self):
        if self.has_end_date_var.get():
            self.end_date_entry.config(state="normal")
        else:
            self.end_date_entry.config(state="disabled")
    
    def on_ok(self):
        try:
            # Validate data
            name = self.name_var.get().strip()
            team = self.team_var.get().strip()
            manager_code = self.manager_code_var.get().strip()
            cost_center = self.cost_center_var.get().strip()
            
            if not name or not manager_code or not cost_center:
                messagebox.showerror("Error", "Name, Manager Code, and Cost Center are required.")
                return
            
            # Parse dates
            try:
                start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Error", "Invalid start date format. Use YYYY-MM-DD.")
                return
            
            end_date = None
            if self.has_end_date_var.get():
                try:
                    end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d").date()
                except ValueError:
                    messagebox.showerror("Error", "Invalid end date format. Use YYYY-MM-DD.")
                    return
            
            employment_type = self.emp_type_var.get()
            
            self.result = {
                "name": name,
                "team": team,
                "manager_code": manager_code,
                "cost_center": cost_center,
                "start_date": start_date,
                "end_date": end_date,
                "employment_type": employment_type
            }
            
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def on_cancel(self):
        self.result = None
        self.destroy()
    
    def get_employee_data(self):
        return self.result

class GA01WeekDialog(tk.Toplevel):
    def __init__(self, parent, year=None):
        super().__init__(parent)
        self.result = None
        
        if year is None:
            year = datetime.now().year
        self.year = year
        
        self.title(f"GA01 Weeks for {year}")
        self.geometry("300x400")
        self.resizable(False, False)
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Month entries
        self.spinboxes = {}
        
        for i, month in enumerate(range(1, 13)):
            month_name = datetime(2000, month, 1).strftime('%B')
            ttk.Label(frame, text=f"{month_name}:").grid(row=i, column=0, sticky=tk.W, pady=5)
            
            var = tk.DoubleVar(value=4.0)
            spinbox = ttk.Spinbox(frame, from_=0, to=5, increment=0.01, textvariable=var, width=10)
            spinbox.grid(row=i, column=1, sticky=tk.W, pady=5)
            self.spinboxes[month] = var
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=12, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Load GA01 weeks from database
        self.load_ga01_weeks()
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def load_ga01_weeks(self):
        """Load GA01 weeks from database"""
        session = get_session()
        ga01_weeks = session.query(GA01Week).filter(GA01Week.year == self.year).all()
        
        for week in ga01_weeks:
            if week.month in self.spinboxes:
                self.spinboxes[week.month].set(week.weeks)
        
        session.close()
    
    def on_ok(self):
        try:
            weeks = {}
            for month, var in self.spinboxes.items():
                weeks[month] = var.get()
            
            self.result = weeks
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def on_cancel(self):
        self.result = None
        self.destroy()
    
    def get_ga01_weeks(self):
        return self.result

class ForecastApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Employee Forecasting Tool")
        self.geometry("1000x600")
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.forecast_tab = ttk.Frame(self.notebook)
        self.employees_tab = ttk.Frame(self.notebook)
        self.ga01_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.forecast_tab, text="Forecast")
        self.notebook.add(self.employees_tab, text="Employees")
        self.notebook.add(self.ga01_tab, text="GA01 Weeks")
        
        # Set up forecast tab
        self.setup_forecast_tab()
        
        # Set up employees tab
        self.setup_employees_tab()
        
        # Set up GA01 tab
        self.setup_ga01_tab()
        
        # Initialize data
        self.load_employees()
        self.load_ga01_weeks()
        self.load_forecast_data()
    
    def setup_forecast_tab(self):
        # Controls frame
        controls_frame = ttk.Frame(self.forecast_tab, padding="5")
        controls_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)
        
        # Year selector
        ttk.Label(controls_frame, text="Year:").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.IntVar(value=datetime.now().year)
        year_spin = ttk.Spinbox(controls_frame, from_=2000, to=2100, textvariable=self.year_var, width=5)
        year_spin.pack(side=tk.LEFT, padx=5)
        
        # Manager filter
        ttk.Label(controls_frame, text="Filter by Manager Code:").pack(side=tk.LEFT, padx=5)
        self.manager_filter_var = tk.StringVar(value="All")
        self.manager_filter = ttk.Combobox(controls_frame, textvariable=self.manager_filter_var, width=15, state="readonly")
        self.manager_filter.pack(side=tk.LEFT, padx=5)
        
        # Cost center filter
        ttk.Label(controls_frame, text="Filter by Cost Center:").pack(side=tk.LEFT, padx=5)
        self.cost_center_filter_var = tk.StringVar(value="All")
        self.cost_center_filter = ttk.Combobox(controls_frame, textvariable=self.cost_center_filter_var, width=15, state="readonly")
        self.cost_center_filter.pack(side=tk.LEFT, padx=5)
        
        # Refresh button
        ttk.Button(controls_frame, text="Refresh", command=self.load_forecast_data).pack(side=tk.LEFT, padx=5)
        
        # Forecast table
        table_frame = ttk.Frame(self.forecast_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        x_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        y_scrollbar = ttk.Scrollbar(table_frame)
        
        # Treeview for forecast data
        self.forecast_tree = ttk.Treeview(table_frame, columns=("manager", "cost_center", "work_code", 
                                                               "jan", "feb", "mar", "apr", "may", "jun",
                                                               "jul", "aug", "sep", "oct", "nov", "dec", "total"),
                                          show="headings", yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Set scrollbars
        x_scrollbar.config(command=self.forecast_tree.xview)
        y_scrollbar.config(command=self.forecast_tree.yview)
        
        # Set column headings
        self.forecast_tree.heading("manager", text="Manager")
        self.forecast_tree.heading("cost_center", text="Cost Center")
        self.forecast_tree.heading("work_code", text="Work Code")
        
        month_abbrs = [datetime(2000, m, 1).strftime('%b') for m in range(1, 13)]
        for i, abbr in enumerate(month_abbrs):
            self.forecast_tree.heading(f"{abbr.lower()}", text=abbr)
        
        self.forecast_tree.heading("total", text="Total")
        
        # Set column widths
        self.forecast_tree.column("manager", width=80)
        self.forecast_tree.column("cost_center", width=80)
        self.forecast_tree.column("work_code", width=100)
        
        for month in ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]:
            self.forecast_tree.column(month, width=60)
        
        self.forecast_tree.column("total", width=80)
        
        # Pack everything
        self.forecast_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Export button
        ttk.Button(self.forecast_tab, text="Export Forecast", command=self.export_forecast).pack(pady=10)
    
    def setup_employees_tab(self):
        # Employees table
        table_frame = ttk.Frame(self.employees_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        x_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        y_scrollbar = ttk.Scrollbar(table_frame)
        
        # Treeview for employees
        self.employees_tree = ttk.Treeview(table_frame, columns=("id", "name", "team", "manager_code", 
                                                                "cost_center", "start_date", "end_date", "type"),
                                          show="headings", yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        
        # Set scrollbars
        x_scrollbar.config(command=self.employees_tree.xview)
        y_scrollbar.config(command=self.employees_tree.yview)
        
        # Set column headings
        self.employees_tree.heading("id", text="ID")
        self.employees_tree.heading("name", text="Name")
        self.employees_tree.heading("team", text="Team")
        self.employees_tree.heading("manager_code", text="Manager Code")
        self.employees_tree.heading("cost_center", text="Cost Center")
        self.employees_tree.heading("start_date", text="Start Date")
        self.employees_tree.heading("end_date", text="End Date")
        self.employees_tree.heading("type", text="Type")
        
        # Set column widths
        self.employees_tree.column("id", width=50)
        self.employees_tree.column("name", width=150)
        self.employees_tree.column("team", width=100)
        self.employees_tree.column("manager_code", width=100)
        self.employees_tree.column("cost_center", width=100)
        self.employees_tree.column("start_date", width=100)
        self.employees_tree.column("end_date", width=100)
        self.employees_tree.column("type", width=100)
        
        # Pack everything
        self.employees_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Controls
        controls_frame = ttk.Frame(self.employees_tab)
        controls_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)
        
        ttk.Button(controls_frame, text="Add Employee", command=self.add_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Edit Employee", command=self.edit_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Remove Employee", command=self.remove_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Import Employees", command=self.import_employees).pack(side=tk.LEFT, padx=5)
    
    def setup_ga01_tab(self):
        # GA01 table
        table_frame = ttk.Frame(self.ga01_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        y_scrollbar = ttk.Scrollbar(table_frame)
        
        # Treeview for GA01 weeks
        self.ga01_tree = ttk.Treeview(table_frame, columns=("year", "month", "weeks"),
                                      show="headings", yscrollcommand=y_scrollbar.set)
        
        # Set scrollbar
        y_scrollbar.config(command=self.ga01_tree.yview)
        
        # Set column headings
        self.ga01_tree.heading("year", text="Year")
        self.ga01_tree.heading("month", text="Month")
        self.ga01_tree.heading("weeks", text="Weeks")
        
        # Set column widths
        self.ga01_tree.column("year", width=100)
        self.ga01_tree.column("month", width=150)
        self.ga01_tree.column("weeks", width=100)
        
        # Pack everything
        self.ga01_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Controls
        controls_frame = ttk.Frame(self.ga01_tab)
        controls_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)
        
        ttk.Button(controls_frame, text="Edit GA01 Weeks", command=self.edit_ga01_weeks).pack(side=tk.LEFT, padx=5)
        ttk.Button(controls_frame, text="Import GA01 Weeks", command=self.import_ga01_weeks).pack(side=tk.LEFT, padx=5)
    
    def load_employees(self):
        """Load employees data into the employees table"""
        # Clear existing data
        for item in self.employees_tree.get_children():
            self.employees_tree.delete(item)
        
        session = get_session()
        employees = session.query(Employee).all()
        
        # Also populate manager code and cost center filters
        manager_codes = {"All"}
        cost_centers = {"All"}
        
        for emp in employees:
            # Format dates
            start_date = emp.start_date.strftime('%Y-%m-%d') if emp.start_date else ""
            end_date = emp.end_date.strftime('%Y-%m-%d') if emp.end_date else ""
            
            # Add to tree
            self.employees_tree.insert("", "end", values=(
                emp.id, emp.name, emp.team or "", emp.manager_code,
                emp.cost_center, start_date, end_date, emp.employment_type
            ))
            
            # Add to filter sets
            manager_codes.add(emp.manager_code)
            cost_centers.add(emp.cost_center)
        
        session.close()
        
        # Update filters while preserving selection if possible
        current_manager = self.manager_filter_var.get()
        self.manager_filter["values"] = sorted(manager_codes)
        if current_manager in manager_codes:
            self.manager_filter_var.set(current_manager)
        else:
            self.manager_filter_var.set("All")
        
        current_cc = self.cost_center_filter_var.get()
        self.cost_center_filter["values"] = sorted(cost_centers)
        if current_cc in cost_centers:
            self.cost_center_filter_var.set(current_cc)
        else:
            self.cost_center_filter_var.set("All")
    
    def load_ga01_weeks(self):
        """Load GA01 weeks data into the GA01 table"""
        # Clear existing data
        for item in self.ga01_tree.get_children():
            self.ga01_tree.delete(item)
        
        session = get_session()
        year = datetime.now().year
        ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
        
        for week in ga01_weeks:
            month_name = datetime(2000, week.month, 1).strftime('%B')
            self.ga01_tree.insert("", "end", values=(
                week.year, month_name, f"{week.weeks:.2f}"
            ))
        
        session.close()
    
    def load_forecast_data(self):
        """Load forecast data into the forecast table"""
        # Clear existing data
        for item in self.forecast_tree.get_children():
            self.forecast_tree.delete(item)
        
        session = get_session()
        year = self.year_var.get()
        
        # Apply filters
        manager_filter = self.manager_filter_var.get()
        cost_center_filter = self.cost_center_filter_var.get()
        
        # Get work codes
        work_codes = session.query(WorkCode).all()
        work_code_dict = {wc.id: wc.code for wc in work_codes}
        
        # Get forecasts for this year
        query = session.query(Forecast).filter(Forecast.year == year)
        
        # Get unique manager/cost center combinations
        if manager_filter != "All" or cost_center_filter != "All":
            # Get valid employee IDs based on filters
            emp_query = session.query(Employee.id)
            
            if manager_filter != "All":
                emp_query = emp_query.filter(Employee.manager_code == manager_filter)
            
            if cost_center_filter != "All":
                emp_query = emp_query.filter(Employee.cost_center == cost_center_filter)
            
            valid_emp_ids = [emp.id for emp in emp_query.all()]
            
            if valid_emp_ids:
                query = query.filter(Forecast.employee_id.in_(valid_emp_ids))
            else:
                # No matching employees, so no forecasts
                session.close()
                return
        
        forecasts = query.all()
        
        # Get unique manager/cost center/work code combinations
        forecast_data = {}
        for forecast in forecasts:
            employee = session.query(Employee).filter(Employee.id == forecast.employee_id).first()
            
            if not employee:
                continue
            
            key = (employee.manager_code, employee.cost_center, forecast.work_code_id)
            
            if key not in forecast_data:
                forecast_data[key] = {
                    "manager_code": employee.manager_code,
                    "cost_center": employee.cost_center,
                    "work_code": work_code_dict.get(forecast.work_code_id, "Unknown"),
                    "months": {}
                }
            
            # Add to monthly total
            if forecast.month not in forecast_data[key]["months"]:
                forecast_data[key]["months"][forecast.month] = 0
            
            forecast_data[key]["months"][forecast.month] += forecast.hours
        
        # Add to tree
        month_columns = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
        
        for key, data in forecast_data.items():
            values = [
                data["manager_code"],
                data["cost_center"],
                data["work_code"]
            ]
            
            total = 0
            for month in range(1, 13):
                hours = data["months"].get(month, 0)
                values.append(f"{hours:.2f}")
                total += hours
            
            values.append(f"{total:.2f}")
            
            self.forecast_tree.insert("", "end", values=values)
        
        session.close()
    
    def add_employee(self):
        """Add a new employee"""
        dialog = EmployeeDialog(self)
        data = dialog.get_employee_data()
        
        if not data:
            return
        
        # Validate data
        if not data["name"] or not data["manager_code"] or not data["cost_center"]:
            messagebox.showerror("Missing Data", "Name, Manager Code, and Cost Center are required.")
            return
        
        # Create new employee
        session = get_session()
        employee = Employee(
            name=data["name"],
            team=data["team"] if data["team"] else None,
            manager_code=data["manager_code"],
            cost_center=data["cost_center"],
            start_date=data["start_date"],
            end_date=data["end_date"],
            employment_type=data["employment_type"]
        )
        
        session.add(employee)
        session.commit()
        session.close()
        
        # Update employee list and forecast
        self.load_employees()
        update_forecast()
        self.load_forecast_data()
    
    def edit_employee(self):
        """Edit the selected employee"""
        selected_item = self.employees_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an employee to edit.")
            return
        
        # Get employee ID from selection
        employee_id = int(self.employees_tree.item(selected_item[0], "values")[0])
        
        session = get_session()
        employee = session.query(Employee).filter(Employee.id == employee_id).first()
        
        if not employee:
            messagebox.showerror("Error", "Employee not found.")
            session.close()
            return
        
        dialog = EmployeeDialog(self, employee)
        data = dialog.get_employee_data()
        
        if not data:
            session.close()
            return
        
        # Validate data
        if not data["name"] or not data["manager_code"] or not data["cost_center"]:
            messagebox.showerror("Missing Data", "Name, Manager Code, and Cost Center are required.")
            session.close()
            return
        
        # Update employee
        employee.name = data["name"]
        employee.team = data["team"] if data["team"] else None
        employee.manager_code = data["manager_code"]
        employee.cost_center = data["cost_center"]
        employee.start_date = data["start_date"]
        employee.end_date = data["end_date"]
        employee.employment_type = data["employment_type"]
        
        session.commit()
        session.close()
        
        # Update employee list and forecast
        self.load_employees()
        update_forecast(employee_id)
        self.load_forecast_data()
    
    def remove_employee(self):
        """Remove the selected employee"""
        selected_item = self.employees_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select an employee to remove.")
            return
        
        # Get employee ID and name from selection
        values = self.employees_tree.item(selected_item[0], "values")
        employee_id = int(values[0])
        employee_name = values[1]
        
        # Confirmation dialog
        if not messagebox.askyesno("Confirm Deletion", 
                                  f"Are you sure you want to delete employee '{employee_name}'?"):
            return
        
        session = get_session()
        
        # Delete forecasts for this employee
        session.query(Forecast).filter(Forecast.employee_id == employee_id).delete()
        
        # Delete employee
        session.query(Employee).filter(Employee.id == employee_id).delete()
        
        session.commit()
        session.close()
        
        # Update employee list and forecast
        self.load_employees()
        self.load_forecast_data()
    
    def edit_ga01_weeks(self):
        """Edit GA01 weeks"""
        year = datetime.now().year
        
        dialog = GA01WeekDialog(self, year)
        weeks_data = dialog.get_ga01_weeks()
        
        if not weeks_data:
            return
        
        session = get_session()
        
        # Delete existing weeks
        session.query(GA01Week).filter(GA01Week.year == year).delete()
        
        # Add new weeks
        for month, weeks in weeks_data.items():
            ga01 = GA01Week(
                year=year,
                month=month,
                weeks=weeks
            )
            session.add(ga01)
        
        session.commit()
        session.close()
        
        # Update GA01 table and forecast
        self.load_ga01_weeks()
        update_forecast()
        self.load_forecast_data()
    
    def import_employees(self):
        """Import employees from file"""
        file_path = filedialog.askopenfilename(
            title="Import Employees",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            import_employees(file_path)
            messagebox.showinfo("Success", "Employees imported successfully.")
            self.load_employees()
            self.load_forecast_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import employees: {str(e)}")
    
    def import_ga01_weeks(self):
        """Import GA01 weeks from file"""
        file_path = filedialog.askopenfilename(
            title="Import GA01 Weeks",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            year = datetime.now().year
            import_ga01_weeks(file_path, year)
            messagebox.showinfo("Success", "GA01 weeks imported successfully.")
            self.load_ga01_weeks()
            self.load_forecast_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import GA01 weeks: {str(e)}")
    
    def export_forecast(self):
        """Export forecast to file"""
        file_path = filedialog.asksaveasfilename(
            title="Export Forecast",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            year = self.year_var.get()
            export_forecast(file_path, year)
            messagebox.showinfo("Success", f"Forecast exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export forecast: {str(e)}")

if __name__ == "__main__":
    app = ForecastApp()
    app.mainloop() 