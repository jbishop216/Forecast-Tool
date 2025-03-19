import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
from models import Employee, GA01Week, WorkCode, Forecast, ProjectAllocation, Settings, PlannedChange, ChangeType, get_session
from forecast_calculator import update_forecast
from import_export import import_employees, import_ga01_weeks, export_forecast
import sqlalchemy

class EmployeeDialog(tk.Toplevel):
    def __init__(self, parent, employee=None):
        super().__init__(parent)
        self.employee = employee
        self.result = None
        
        self.title("Employee Details")
        self.geometry("550x450")
        self.resizable(True, True)
        
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
        self.geometry("450x550")
        self.resizable(True, True)
        
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

class ProjectAllocationDialog(tk.Toplevel):
    def __init__(self, parent, manager_code, year=None):
        super().__init__(parent)
        self.result = None
        self.manager_code = manager_code
        
        if year is None:
            year = datetime.now().year
        self.year = year
        
        self.title(f"Project Allocations for {manager_code} - {year}")
        self.geometry("450x600")
        self.resizable(True, True)
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Paste functionality
        paste_frame = ttk.Frame(frame)
        paste_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        ttk.Label(paste_frame, text="Excel Data:").pack(side=tk.LEFT, padx=5)
        ttk.Button(paste_frame, text="Paste", command=self.paste_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(paste_frame, text="Copy", command=self.copy_data).pack(side=tk.LEFT, padx=5)
        
        # Instructions
        ttk.Label(frame, text="Format: 12 numbers (Jan-Dec)", 
                  font=("", 8)).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Month entries
        self.spinboxes = {}
        
        for i, month in enumerate(range(1, 13)):
            month_name = datetime(2000, month, 1).strftime('%B')
            ttk.Label(frame, text=f"{month_name}:").grid(row=i+2, column=0, sticky=tk.W, pady=5)
            
            var = tk.DoubleVar(value=0.0)
            spinbox = ttk.Spinbox(frame, from_=0, to=1000, increment=0.5, textvariable=var, width=10)
            spinbox.grid(row=i+2, column=1, sticky=tk.W, pady=5)
            self.spinboxes[month] = var
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=14, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Load existing allocations
        self.load_allocations()
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def paste_data(self):
        """Handle pasting data from clipboard"""
        try:
            # Get clipboard data
            clipboard_data = self.clipboard_get()
            
            # Process clipboard data - try different delimiters
            values = []
            
            # Try tab-delimited (Excel default)
            if '\t' in clipboard_data:
                values = clipboard_data.strip().split('\t')
            # Try comma-delimited
            elif ',' in clipboard_data:
                values = clipboard_data.strip().split(',')
            # Try space-delimited
            elif ' ' in clipboard_data:
                values = clipboard_data.strip().split()
            # Try newline-delimited
            elif '\n' in clipboard_data:
                values = clipboard_data.strip().split('\n')
            else:
                # Just a single value
                values = [clipboard_data.strip()]
            
            # Convert values to numbers
            numeric_values = []
            for val in values:
                try:
                    numeric_values.append(float(val.strip()))
                except ValueError:
                    pass  # Skip non-numeric values
            
            # Apply values to the spinboxes
            for month, value in enumerate(numeric_values[:12], start=1):
                if month in self.spinboxes:
                    self.spinboxes[month].set(value)
            
            messagebox.showinfo("Paste Complete", f"Successfully pasted {len(numeric_values[:12])} values.")
            
        except Exception as e:
            messagebox.showerror("Paste Error", f"Error pasting data: {str(e)}")
    
    def copy_data(self):
        """Copy current values to clipboard in Excel-friendly format"""
        try:
            # Get values from spinboxes
            values = []
            for month in range(1, 13):
                if month in self.spinboxes:
                    values.append(str(self.spinboxes[month].get()))
            
            # Create tab-delimited string (Excel format)
            clipboard_data = "\t".join(values)
            
            # Copy to clipboard
            self.clipboard_clear()
            self.clipboard_append(clipboard_data)
            
            messagebox.showinfo("Copy Complete", "Current values copied to clipboard.")
        
        except Exception as e:
            messagebox.showerror("Copy Error", f"Error copying data: {str(e)}")
    
    def load_allocations(self):
        """Load project allocations from database"""
        session = get_session()
        allocations = session.query(ProjectAllocation).filter(
            ProjectAllocation.manager_code == self.manager_code,
            ProjectAllocation.year == self.year
        ).all()
        
        for allocation in allocations:
            if allocation.month in self.spinboxes:
                self.spinboxes[allocation.month].set(allocation.hours)
        
        session.close()
    
    def on_ok(self):
        try:
            hours = {}
            for month, var in self.spinboxes.items():
                hours[month] = var.get()
            
            self.result = hours
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def on_cancel(self):
        self.result = None
        self.destroy()
    
    def get_allocations(self):
        return self.result

class SettingsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.result = None
        
        self.title("Hour Settings")
        self.geometry("450x250")
        self.resizable(True, True)
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="FTE Weekly Hours:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.fte_hours_var = tk.DoubleVar(value=34.5)
        ttk.Spinbox(frame, from_=0, to=100, increment=0.5, textvariable=self.fte_hours_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(frame, text="Contractor Weekly Hours:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.contractor_hours_var = tk.DoubleVar(value=39.0)
        ttk.Spinbox(frame, from_=0, to=100, increment=0.5, textvariable=self.contractor_hours_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Load current settings
        self.load_settings()
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def load_settings(self):
        """Load settings from database"""
        session = get_session()
        settings = session.query(Settings).first()
        
        if settings:
            self.fte_hours_var.set(settings.fte_hours)
            self.contractor_hours_var.set(settings.contractor_hours)
        
        session.close()
    
    def on_ok(self):
        try:
            self.result = {
                "fte_hours": self.fte_hours_var.get(),
                "contractor_hours": self.contractor_hours_var.get()
            }
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def on_cancel(self):
        self.result = None
        self.destroy()
    
    def get_settings(self):
        return self.result

class PlannedChangeDialog(tk.Toplevel):
    def __init__(self, parent, planned_change=None):
        super().__init__(parent)
        self.planned_change = planned_change
        self.result = None
        self.employees = {}
        
        self.title("Planned Employee Change")
        self.geometry("650x550")
        self.resizable(True, True)
        
        # Load employees for reference
        session = get_session()
        employees = session.query(Employee).all()
        for emp in employees:
            self.employees[emp.id] = emp
        session.close()
        
        # Create form fields
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Description
        ttk.Label(frame, text="Description:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.description_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.description_var, width=40).grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        # Change Type
        ttk.Label(frame, text="Change Type:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.change_type_var = tk.StringVar()
        change_types = [ct.value for ct in ChangeType]
        self.change_type_combo = ttk.Combobox(frame, textvariable=self.change_type_var, 
                                             values=change_types, width=20, state="readonly")
        self.change_type_combo.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=5)
        self.change_type_combo.bind("<<ComboboxSelected>>", self.on_change_type_changed)
        
        # Effective Date
        ttk.Label(frame, text="Effective Date (YYYY-MM-DD):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.effective_date_var = tk.StringVar()
        self.effective_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(frame, textvariable=self.effective_date_var, width=20).grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        # Employee Selection Frame (for Conversion and Termination)
        self.employee_frame = ttk.LabelFrame(frame, text="Select Employee", padding="10")
        self.employee_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        ttk.Label(self.employee_frame, text="Employee:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.employee_id_var = tk.StringVar()
        self.employee_combo = ttk.Combobox(self.employee_frame, textvariable=self.employee_id_var, width=40, state="readonly")
        self.employee_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Populate employee combo
        employee_options = [""]
        self.employee_id_map = {"": None}
        for emp_id, emp in self.employees.items():
            display_text = f"{emp.name} ({emp.manager_code}, {emp.employment_type})"
            employee_options.append(display_text)
            self.employee_id_map[display_text] = emp_id
        
        self.employee_combo["values"] = employee_options
        
        # New Hire Frame
        self.new_hire_frame = ttk.LabelFrame(frame, text="New Hire Details", padding="10")
        self.new_hire_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        ttk.Label(self.new_hire_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        ttk.Entry(self.new_hire_frame, textvariable=self.name_var, width=30).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.new_hire_frame, text="Team:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.team_var = tk.StringVar()
        ttk.Entry(self.new_hire_frame, textvariable=self.team_var, width=30).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.new_hire_frame, text="Manager Code:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.manager_code_var = tk.StringVar()
        ttk.Entry(self.new_hire_frame, textvariable=self.manager_code_var, width=30).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.new_hire_frame, text="Cost Center:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(self.new_hire_frame, textvariable=self.cost_center_var, width=30).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.new_hire_frame, text="Employment Type:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.employment_type_var = tk.StringVar(value="FTE")
        ttk.Combobox(self.new_hire_frame, textvariable=self.employment_type_var, values=["FTE", "Contractor"], 
                     width=20, state="readonly").grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # Conversion Frame
        self.conversion_frame = ttk.LabelFrame(frame, text="Conversion Details", padding="10")
        self.conversion_frame.grid(row=5, column=0, columnspan=3, sticky=tk.W+tk.E, pady=10)
        
        ttk.Label(self.conversion_frame, text="Target Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.target_type_var = tk.StringVar(value="FTE")
        ttk.Combobox(self.conversion_frame, textvariable=self.target_type_var, values=["FTE", "Contractor"], 
                     width=20, state="readonly").grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Initialize state based on change type
        self.change_type_var.set(ChangeType.NEW_HIRE.value)
        self.on_change_type_changed(None)
        
        # Populate fields if editing existing planned change
        if planned_change:
            self.description_var.set(planned_change.description)
            self.change_type_var.set(planned_change.change_type)
            
            if planned_change.effective_date:
                self.effective_date_var.set(planned_change.effective_date.strftime("%Y-%m-%d"))
            
            if planned_change.employee_id:
                for display_text, emp_id in self.employee_id_map.items():
                    if emp_id == planned_change.employee_id:
                        self.employee_id_var.set(display_text)
                        break
            
            # New hire fields
            if planned_change.name:
                self.name_var.set(planned_change.name)
            if planned_change.team:
                self.team_var.set(planned_change.team)
            if planned_change.manager_code:
                self.manager_code_var.set(planned_change.manager_code)
            if planned_change.cost_center:
                self.cost_center_var.set(planned_change.cost_center)
            if planned_change.employment_type:
                self.employment_type_var.set(planned_change.employment_type)
            
            # Conversion fields
            if planned_change.target_employment_type:
                self.target_type_var.set(planned_change.target_employment_type)
            
            # Update UI state
            self.on_change_type_changed(None)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def on_change_type_changed(self, event):
        """Update UI based on selected change type"""
        change_type = self.change_type_var.get()
        
        # Hide all specialized frames first
        self.employee_frame.grid_remove()
        self.new_hire_frame.grid_remove()
        self.conversion_frame.grid_remove()
        
        if change_type == ChangeType.NEW_HIRE.value:
            self.new_hire_frame.grid()
        elif change_type == ChangeType.CONVERSION.value:
            self.employee_frame.grid()
            self.conversion_frame.grid()
        elif change_type == ChangeType.TERMINATION.value:
            self.employee_frame.grid()
    
    def on_ok(self):
        try:
            # Validate data
            description = self.description_var.get().strip()
            change_type = self.change_type_var.get()
            
            if not description or not change_type:
                messagebox.showerror("Error", "Description and Change Type are required.")
                return
            
            # Parse effective date
            try:
                effective_date = datetime.strptime(self.effective_date_var.get(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Error", "Invalid effective date format. Use YYYY-MM-DD.")
                return
            
            # Get employee ID if applicable
            employee_id = None
            if change_type in [ChangeType.CONVERSION.value, ChangeType.TERMINATION.value]:
                employee_display = self.employee_id_var.get()
                if not employee_display:
                    messagebox.showerror("Error", "Please select an employee.")
                    return
                employee_id = self.employee_id_map.get(employee_display)
            
            # Validate new hire details
            name = team = manager_code = cost_center = employment_type = None
            if change_type == ChangeType.NEW_HIRE.value:
                name = self.name_var.get().strip()
                team = self.team_var.get().strip()
                manager_code = self.manager_code_var.get().strip()
                cost_center = self.cost_center_var.get().strip()
                employment_type = self.employment_type_var.get()
                
                if not name or not manager_code or not cost_center:
                    messagebox.showerror("Error", "Name, Manager Code, and Cost Center are required for new hires.")
                    return
            
            # Validate conversion details
            target_employment_type = None
            if change_type == ChangeType.CONVERSION.value:
                target_employment_type = self.target_type_var.get()
                
                # Check if conversion would be a no-op
                if employee_id and target_employment_type:
                    employee = self.employees.get(employee_id)
                    if employee and employee.employment_type == target_employment_type:
                        messagebox.showerror("Error", f"Employee is already a {target_employment_type}.")
                        return
            
            self.result = {
                "description": description,
                "change_type": change_type,
                "effective_date": effective_date,
                "employee_id": employee_id,
                "name": name,
                "team": team,
                "manager_code": manager_code,
                "cost_center": cost_center,
                "employment_type": employment_type,
                "target_employment_type": target_employment_type
            }
            
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def on_cancel(self):
        self.result = None
        self.destroy()
    
    def get_planned_change_data(self):
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
        self.settings_tab = ttk.Frame(self.notebook)
        self.planned_changes_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.forecast_tab, text="Forecast")
        self.notebook.add(self.employees_tab, text="Employees")
        self.notebook.add(self.ga01_tab, text="GA01 Weeks")
        self.notebook.add(self.planned_changes_tab, text="Planned Changes")
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Set up forecast tab
        self.setup_forecast_tab()
        
        # Set up employees tab
        self.setup_employees_tab()
        
        # Set up GA01 tab
        self.setup_ga01_tab()
        
        # Set up settings tab
        self.setup_settings_tab()
        
        # Set up planned changes tab
        self.setup_planned_changes_tab()
        
        # Initialize data
        self.load_employees()
        self.load_ga01_weeks()
        self.load_forecast_data()
        self.initialize_settings()
        self.load_planned_changes()
    
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
        
        # Project Allocation button
        ttk.Button(controls_frame, text="Edit Project Allocation", command=self.edit_project_allocation).pack(side=tk.LEFT, padx=5)
        
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
    
    def setup_settings_tab(self):
        """Set up the settings tab"""
        frame = ttk.Frame(self.settings_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Hours settings frame
        hours_frame = ttk.LabelFrame(frame, text="Weekly Hours Settings", padding="10")
        hours_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)
        
        # FTE hours
        ttk.Label(hours_frame, text="FTE Weekly Hours:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.fte_hours_var = tk.DoubleVar(value=34.5)
        ttk.Spinbox(hours_frame, from_=0, to=100, increment=0.5, textvariable=self.fte_hours_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Contractor hours
        ttk.Label(hours_frame, text="Contractor Weekly Hours:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.contractor_hours_var = tk.DoubleVar(value=39.0)
        ttk.Spinbox(hours_frame, from_=0, to=100, increment=0.5, textvariable=self.contractor_hours_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Save button
        ttk.Button(hours_frame, text="Save Hours Settings", command=self.save_hour_settings).grid(row=2, column=0, columnspan=2, pady=10)
    
    def setup_planned_changes_tab(self):
        """Set up the planned changes tab"""
        # Controls frame
        controls_frame = ttk.Frame(self.planned_changes_tab, padding="5")
        controls_frame.pack(fill=tk.X, expand=False, padx=5, pady=5)
        
        # Add planned change button
        ttk.Button(controls_frame, text="Add Planned Change", 
                  command=self.add_planned_change).pack(side=tk.LEFT, padx=5)
        
        # Edit planned change button
        ttk.Button(controls_frame, text="Edit Planned Change", 
                  command=self.edit_planned_change).pack(side=tk.LEFT, padx=5)
        
        # Remove planned change button
        ttk.Button(controls_frame, text="Remove Planned Change", 
                  command=self.remove_planned_change).pack(side=tk.LEFT, padx=5)
        
        # Apply changes button
        ttk.Button(controls_frame, text="Apply Changes to Forecast", 
                  command=self.apply_planned_changes).pack(side=tk.LEFT, padx=20)
        
        # Table frame
        table_frame = ttk.Frame(self.planned_changes_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        x_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        y_scrollbar = ttk.Scrollbar(table_frame)
        
        # Treeview for planned changes
        self.planned_changes_tree = ttk.Treeview(table_frame, 
                                               columns=("id", "description", "type", "date", "employee", "details"),
                                               show="headings", 
                                               yscrollcommand=y_scrollbar.set, 
                                               xscrollcommand=x_scrollbar.set)
        
        # Set scrollbars
        x_scrollbar.config(command=self.planned_changes_tree.xview)
        y_scrollbar.config(command=self.planned_changes_tree.yview)
        
        # Set column headings
        self.planned_changes_tree.heading("id", text="ID")
        self.planned_changes_tree.heading("description", text="Description")
        self.planned_changes_tree.heading("type", text="Change Type")
        self.planned_changes_tree.heading("date", text="Effective Date")
        self.planned_changes_tree.heading("employee", text="Employee")
        self.planned_changes_tree.heading("details", text="Details")
        
        # Set column widths
        self.planned_changes_tree.column("id", width=40)
        self.planned_changes_tree.column("description", width=150)
        self.planned_changes_tree.column("type", width=100)
        self.planned_changes_tree.column("date", width=100)
        self.planned_changes_tree.column("employee", width=150)
        self.planned_changes_tree.column("details", width=250)
        
        # Pack everything
        self.planned_changes_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def load_planned_changes(self):
        """Load planned changes into the treeview"""
        # Clear existing data
        for item in self.planned_changes_tree.get_children():
            self.planned_changes_tree.delete(item)
        
        session = get_session()
        
        # Get all planned changes sorted by effective date
        planned_changes = session.query(PlannedChange).order_by(PlannedChange.effective_date).all()
        
        employee_cache = {}
        
        for change in planned_changes:
            # Get employee name if applicable
            employee_name = ""
            if change.employee_id:
                if change.employee_id in employee_cache:
                    employee_name = employee_cache[change.employee_id]
                else:
                    employee = session.query(Employee).filter(Employee.id == change.employee_id).first()
                    if employee:
                        employee_name = employee.name
                        employee_cache[change.employee_id] = employee_name
            
            # Format details based on change type
            details = ""
            if change.change_type == ChangeType.NEW_HIRE.value:
                details = f"{change.name} ({change.manager_code}, {change.cost_center}) as {change.employment_type}"
            elif change.change_type == ChangeType.CONVERSION.value:
                details = f"Convert to {change.target_employment_type}"
            elif change.change_type == ChangeType.TERMINATION.value:
                details = "Termination"
            
            # Format date
            date_str = change.effective_date.strftime('%Y-%m-%d') if change.effective_date else ""
            
            # Add to tree
            self.planned_changes_tree.insert("", "end", values=(
                change.id, change.description, change.change_type, date_str, employee_name, details
            ))
        
        session.close()
    
    def add_planned_change(self):
        """Add a new planned change"""
        dialog = PlannedChangeDialog(self)
        data = dialog.get_planned_change_data()
        
        if not data:
            return
        
        session = get_session()
        
        # Create new planned change
        planned_change = PlannedChange(
            description=data["description"],
            change_type=data["change_type"],
            effective_date=data["effective_date"],
            employee_id=data["employee_id"],
            name=data["name"],
            team=data["team"],
            manager_code=data["manager_code"],
            cost_center=data["cost_center"],
            employment_type=data["employment_type"],
            target_employment_type=data["target_employment_type"]
        )
        
        session.add(planned_change)
        session.commit()
        session.close()
        
        # Refresh planned changes list
        self.load_planned_changes()
    
    def edit_planned_change(self):
        """Edit the selected planned change"""
        selected_item = self.planned_changes_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a planned change to edit.")
            return
        
        # Get planned change ID from selection
        planned_change_id = int(self.planned_changes_tree.item(selected_item[0], "values")[0])
        
        session = get_session()
        planned_change = session.query(PlannedChange).filter(PlannedChange.id == planned_change_id).first()
        
        if not planned_change:
            messagebox.showerror("Error", "Planned change not found.")
            session.close()
            return
        
        # Show dialog with planned change data
        dialog = PlannedChangeDialog(self, planned_change)
        data = dialog.get_planned_change_data()
        
        if not data:
            session.close()
            return
        
        # Update planned change
        planned_change.description = data["description"]
        planned_change.change_type = data["change_type"]
        planned_change.effective_date = data["effective_date"]
        planned_change.employee_id = data["employee_id"]
        planned_change.name = data["name"]
        planned_change.team = data["team"]
        planned_change.manager_code = data["manager_code"]
        planned_change.cost_center = data["cost_center"]
        planned_change.employment_type = data["employment_type"]
        planned_change.target_employment_type = data["target_employment_type"]
        
        session.commit()
        session.close()
        
        # Refresh planned changes list
        self.load_planned_changes()
    
    def remove_planned_change(self):
        """Remove the selected planned change"""
        selected_item = self.planned_changes_tree.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a planned change to remove.")
            return
        
        # Get planned change ID and description from selection
        values = self.planned_changes_tree.item(selected_item[0], "values")
        planned_change_id = int(values[0])
        description = values[1]
        
        # Confirmation dialog
        if not messagebox.askyesno("Confirm Deletion", 
                                  f"Are you sure you want to delete planned change '{description}'?"):
            return
        
        session = get_session()
        
        # Delete planned change
        session.query(PlannedChange).filter(PlannedChange.id == planned_change_id).delete()
        
        session.commit()
        session.close()
        
        # Refresh planned changes list
        self.load_planned_changes()
    
    def apply_planned_changes(self):
        """Apply planned changes to create actual employees"""
        if not messagebox.askyesno("Confirm Apply", 
                                 "Are you sure you want to apply all planned changes?\n"
                                 "This will make actual changes to your employee data."):
            return
        
        session = get_session()
        
        # Get all planned changes ordered by date
        planned_changes = session.query(PlannedChange).order_by(PlannedChange.effective_date).all()
        
        changes_applied = 0
        current_date = datetime.now().date()
        
        for change in planned_changes:
            # Only apply changes with effective dates in the past
            if change.effective_date > current_date:
                continue
            
            if change.change_type == ChangeType.NEW_HIRE.value:
                # Create new employee
                employee = Employee(
                    name=change.name,
                    team=change.team,
                    manager_code=change.manager_code,
                    cost_center=change.cost_center,
                    start_date=change.effective_date,
                    employment_type=change.employment_type
                )
                session.add(employee)
                changes_applied += 1
                
            elif change.change_type == ChangeType.CONVERSION.value:
                # Convert existing employee's type
                employee = session.query(Employee).filter(Employee.id == change.employee_id).first()
                if employee:
                    employee.employment_type = change.target_employment_type
                    changes_applied += 1
                
            elif change.change_type == ChangeType.TERMINATION.value:
                # Set end date for employee
                employee = session.query(Employee).filter(Employee.id == change.employee_id).first()
                if employee:
                    employee.end_date = change.effective_date
                    changes_applied += 1
            
            # Delete the planned change after applying
            session.delete(change)
        
        session.commit()
        session.close()
        
        # Update the data
        self.load_employees()
        
        # Update the forecast with the latest employee changes
        update_forecast()
        
        # Refresh the UI
        self.load_forecast_data()
        self.load_planned_changes()
        
        messagebox.showinfo("Changes Applied", f"{changes_applied} changes were applied successfully.")
    
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
    
    def initialize_settings(self):
        """Initialize settings from database"""
        session = get_session()
        
        # Check if settings exist, if not create default
        settings = session.query(Settings).first()
        if not settings:
            settings = Settings(fte_hours=34.5, contractor_hours=39.0)
            session.add(settings)
            session.commit()
        
        # Update UI
        self.fte_hours_var.set(settings.fte_hours)
        self.contractor_hours_var.set(settings.contractor_hours)
        
        session.close()
    
    def save_hour_settings(self):
        """Save hour settings to database"""
        fte_hours = self.fte_hours_var.get()
        contractor_hours = self.contractor_hours_var.get()
        
        if fte_hours <= 0 or contractor_hours <= 0:
            messagebox.showerror("Invalid Hours", "Hours must be greater than zero.")
            return
        
        session = get_session()
        
        # Get or create settings
        settings = session.query(Settings).first()
        if not settings:
            settings = Settings()
            session.add(settings)
        
        # Update settings
        settings.fte_hours = fte_hours
        settings.contractor_hours = contractor_hours
        
        session.commit()
        session.close()
        
        messagebox.showinfo("Success", "Hour settings saved successfully.")
        
        # Update forecast with new hours
        update_forecast()
        self.load_forecast_data()
    
    def edit_project_allocation(self):
        """Edit project allocation for a manager code"""
        # Check if a manager is selected in the filter
        manager_code = self.manager_filter_var.get()
        if manager_code == "All":
            # If no manager is selected, ask the user to select one
            session = get_session()
            manager_codes = [code for code in self.manager_filter["values"] if code != "All"]
            session.close()
            
            if not manager_codes:
                messagebox.showwarning("No Managers", "No manager codes found. Please add employees first.")
                return
            
            # Create a dialog to select a manager code
            select_dialog = tk.Toplevel(self)
            select_dialog.title("Select Manager Code")
            select_dialog.geometry("350x200")
            select_dialog.resizable(True, True)
            select_dialog.transient(self)
            select_dialog.grab_set()
            
            frame = ttk.Frame(select_dialog, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(frame, text="Select Manager Code:").pack(pady=5)
            
            selected_manager = tk.StringVar()
            manager_combo = ttk.Combobox(frame, textvariable=selected_manager, values=manager_codes, width=20, state="readonly")
            manager_combo.pack(pady=5)
            if manager_codes:
                manager_combo.current(0)
            
            def on_ok():
                nonlocal manager_code
                manager_code = selected_manager.get()
                select_dialog.destroy()
                if manager_code:
                    self.show_project_allocation_dialog(manager_code)
            
            def on_cancel():
                select_dialog.destroy()
            
            button_frame = ttk.Frame(frame)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="Cancel", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)
            
            self.wait_window(select_dialog)
        else:
            # If a manager is already selected, use that
            self.show_project_allocation_dialog(manager_code)
    
    def show_project_allocation_dialog(self, manager_code):
        """Show project allocation dialog for the specified manager code"""
        year = self.year_var.get()
        
        dialog = ProjectAllocationDialog(self, manager_code, year)
        allocation_data = dialog.get_allocations()
        
        if not allocation_data:
            return
        
        session = get_session()
        
        # Delete existing allocations for this manager and year
        session.query(ProjectAllocation).filter(
            ProjectAllocation.manager_code == manager_code,
            ProjectAllocation.year == year
        ).delete()
        
        # Add new allocations
        for month, hours in allocation_data.items():
            allocation = ProjectAllocation(
                manager_code=manager_code,
                year=year,
                month=month,
                hours=hours
            )
            session.add(allocation)
        
        session.commit()
        session.close()
        
        # Update forecast with new allocations
        update_forecast()
        self.load_forecast_data()
        
        messagebox.showinfo("Success", f"Project allocations for {manager_code} saved successfully.")
    
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
                # For employee-based forecasts
                query_conditions = [(Forecast.employee_id.in_(valid_emp_ids))]
                
                # Also include planned hires that match the filters
                if manager_filter != "All" and cost_center_filter == "All":
                    # When only manager filter is active
                    query_conditions.append(
                        (Forecast.employee_id.is_(None) & 
                         Forecast.notes.contains(f"({manager_filter},"))
                    )
                elif manager_filter == "All" and cost_center_filter != "All":
                    # When only cost center filter is active
                    query_conditions.append(
                        (Forecast.employee_id.is_(None) & 
                         Forecast.notes.contains(f", {cost_center_filter})"))
                    )
                elif manager_filter != "All" and cost_center_filter != "All":
                    # When both filters are active
                    query_conditions.append(
                        (Forecast.employee_id.is_(None) & 
                         Forecast.notes.contains(f"({manager_filter}, {cost_center_filter})"))
                    )
                else:
                    # When no filters are active, include all planned hires
                    query_conditions.append(
                        (Forecast.employee_id.is_(None) & Forecast.notes.isnot(None))
                    )
                
                query = query.filter(sqlalchemy.or_(*query_conditions))
            else:
                # No matching employees based on filters, but still check for planned hires
                if manager_filter != "All":
                    query = query.filter(
                        (Forecast.employee_id.is_(None) & 
                         Forecast.notes.contains(f"({manager_filter},"))
                    )
                elif cost_center_filter != "All":
                    query = query.filter(
                        (Forecast.employee_id.is_(None) & 
                         Forecast.notes.contains(f", {cost_center_filter})"))
                    )
                else:
                    # No matching criteria, return no results
                    session.close()
                    return
        
        forecasts = query.all()
        
        # Get unique manager/cost center/work code combinations
        forecast_data = {}
        for forecast in forecasts:
            # Handle regular employees
            if forecast.employee_id:
                employee = session.query(Employee).filter(Employee.id == forecast.employee_id).first()
                
                if not employee:
                    continue
                
                manager_code = employee.manager_code
                cost_center = employee.cost_center
            # Handle planned hires (with notes but no employee_id)
            elif forecast.notes:
                try:
                    # Parse the manager_code and cost_center from the notes
                    # Format: "Future hire: Name (manager_code, cost_center) planned date"
                    if '(' in forecast.notes and ')' in forecast.notes:
                        # Extract the content between parentheses
                        paren_content = forecast.notes.split('(')[1].split(')')[0]
                        # Split by comma and trim whitespace
                        parts = [p.strip() for p in paren_content.split(',')]
                        if len(parts) >= 2:
                            manager_code = parts[0]
                            cost_center = parts[1]
                        else:
                            manager_code = "Planned"
                            cost_center = "Future hires"
                    else:
                        manager_code = "Planned"
                        cost_center = "Future hires"
                except Exception:
                    # Fallback if parsing fails
                    manager_code = "Planned"
                    cost_center = "Future hires"
            else:
                continue
            
            key = (manager_code, cost_center, forecast.work_code_id)
            
            if key not in forecast_data:
                forecast_data[key] = {
                    "manager_code": manager_code,
                    "cost_center": cost_center,
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

if __name__ == "__main__":
    app = ForecastApp()
    app.mainloop() 