import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, Enum
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import enum
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# Create database engine
engine = create_engine('sqlite:///forecast.db')
Base = declarative_base()
Session = sessionmaker(bind=engine)

def get_session():
    return Session()

# Enums
class EmploymentType(enum.Enum):
    FTE = "FTE"
    CONTRACTOR = "CONTRACTOR"

class ChangeType(enum.Enum):
    NEW_HIRE = "New Hire"
    CONVERSION = "Conversion"
    TERMINATION = "Termination"

# Database Models
class Employee(Base):
    __tablename__ = 'employees'
    
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    manager_code = Column(String, nullable=False)
    employment_type = Column(String, nullable=False)
    start_date = Column(Date, nullable=False)
    end_date = Column(Date)

class ProjectAllocation(Base):
    __tablename__ = 'project_allocations'
    
    id = Column(Integer, primary_key=True)
    manager_code = Column(String, nullable=False)
    cost_center = Column(String, nullable=False)
    work_code_id = Column(String, nullable=False)
    year = Column(Integer, nullable=False)
    jan = Column(Float, default=0)
    feb = Column(Float, default=0)
    mar = Column(Float, default=0)
    apr = Column(Float, default=0)
    may = Column(Float, default=0)
    jun = Column(Float, default=0)
    jul = Column(Float, default=0)
    aug = Column(Float, default=0)
    sep = Column(Float, default=0)
    oct = Column(Float, default=0)
    nov = Column(Float, default=0)
    dec = Column(Float, default=0)

class Settings(Base):
    __tablename__ = 'settings'
    
    id = Column(Integer, primary_key=True)
    fte_hours = Column(Float, default=34.5)
    contractor_hours = Column(Float, default=39.0)

class GA01Week(Base):
    __tablename__ = 'ga01_weeks'
    
    id = Column(Integer, primary_key=True)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    weeks = Column(Float, nullable=False)

class PlannedChange(Base):
    __tablename__ = 'planned_changes'
    
    id = Column(Integer, primary_key=True)
    description = Column(String, nullable=False)
    change_type = Column(String, nullable=False)
    effective_date = Column(Date, nullable=False)
    employee_id = Column(Integer)
    target_type = Column(String)
    name = Column(String)
    team = Column(String)
    manager_code = Column(String)
    cost_center = Column(String)
    employment_type = Column(String)
    status = Column(String, nullable=False)

class Forecast(Base):
    __tablename__ = 'forecasts'
    
    id = Column(Integer, primary_key=True)
    year = Column(Integer, nullable=False)
    cost_center = Column(String, nullable=False)
    manager_code = Column(String, nullable=False)
    work_code_id = Column(String, nullable=False)
    jan = Column(Float, default=0)
    feb = Column(Float, default=0)
    mar = Column(Float, default=0)
    apr = Column(Float, default=0)
    may = Column(Float, default=0)
    jun = Column(Float, default=0)
    jul = Column(Float, default=0)
    aug = Column(Float, default=0)
    sep = Column(Float, default=0)
    oct = Column(Float, default=0)
    nov = Column(Float, default=0)
    dec = Column(Float, default=0)
    total_hours = Column(Float, default=0)

# Create tables
Base.metadata.create_all(engine)

class EmployeeDialog(tk.Toplevel):
    def __init__(self, parent, employee=None):
        super().__init__(parent)
        self.employee = employee
        self.result = None
        
        self.title("Employee")
        self.geometry("450x350")
        self.resizable(True, True)
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        ttk.Label(frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Manager Code
        ttk.Label(frame, text="Manager Code:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.manager_code_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.manager_code_var, width=20).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Employment Type
        ttk.Label(frame, text="Employment Type:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.employment_type_var = tk.StringVar()
        employment_types = [et.value for et in EmploymentType]
        ttk.Combobox(frame, textvariable=self.employment_type_var, values=employment_types, width=20, state="readonly").grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Start Date
        ttk.Label(frame, text="Start Date (YYYY-MM-DD):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.start_date_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.start_date_var, width=20).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # End Date
        ttk.Label(frame, text="End Date (YYYY-MM-DD):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.end_date_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.end_date_var, width=20).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # If editing, populate fields
        if employee:
            self.name_var.set(employee.name)
            self.manager_code_var.set(employee.manager_code)
            self.employment_type_var.set(employee.employment_type)
            self.start_date_var.set(employee.start_date.strftime("%Y-%m-%d") if employee.start_date else "")
            self.end_date_var.set(employee.end_date.strftime("%Y-%m-%d") if employee.end_date else "")
        else:
            self.start_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def on_ok(self):
        try:
            # Validate required fields
            if not self.name_var.get().strip():
                raise ValueError("Name is required")
            if not self.manager_code_var.get().strip():
                raise ValueError("Manager Code is required")
            if not self.employment_type_var.get():
                raise ValueError("Employment Type is required")
            if not self.start_date_var.get().strip():
                raise ValueError("Start Date is required")
            
            # Parse dates
            start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d").date()
            end_date = None
            if self.end_date_var.get().strip():
                end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d").date()
                if end_date <= start_date:
                    raise ValueError("End Date must be after Start Date")
            
            self.result = {
                "name": self.name_var.get().strip(),
                "manager_code": self.manager_code_var.get().strip(),
                "employment_type": self.employment_type_var.get(),
                "start_date": start_date,
                "end_date": end_date
            }
            self.destroy()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
    
    def on_cancel(self):
        self.result = None
        self.destroy()

class EmployeeTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="10")
        self.parent = parent
        self.create_widgets()
        
    def create_widgets(self):
        # Create toolbar
        toolbar = ttk.Frame(self, style="Panel.TFrame")
        toolbar.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Add decorative toolbar accent
        accent_canvas = tk.Canvas(toolbar, width=3, height=30, highlightthickness=0)
        accent_canvas.pack(side=tk.LEFT, padx=(0, 10), fill=tk.Y)
        # Draw a gradient accent
        for i in range(30):
            # Gradient from primary to secondary 
            r = int(74 + (52-74) * i/30)
            g = int(125 + (195-125) * i/30)
            b = int(186 + (143-186) * i/30)
            color = f'#{r:02x}{g:02x}{b:02x}'
            accent_canvas.create_line(0, i, 3, i, fill=color)
        
        # Button container for better visual grouping
        button_container = ttk.Frame(toolbar, style="Panel.TFrame")
        button_container.pack(side=tk.LEFT, fill=tk.Y)
        
        ttk.Button(button_container, text="Add Employee", command=self.add_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Edit Employee", command=self.edit_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Delete Employee", command=self.delete_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Refresh", command=self.load_employees).pack(side=tk.LEFT, padx=5)
        
        # Create treeview
        self.tree = ttk.Treeview(self, columns=("ID", "Name", "Manager Code", "Type", "Start Date", "End Date"), show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure treeview columns
        self.tree.heading("ID", text="ID")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Manager Code", text="Manager Code")
        self.tree.heading("Type", text="Type")
        self.tree.heading("Start Date", text="Start Date")
        self.tree.heading("End Date", text="End Date")
        
        # Configure column widths
        self.tree.column("ID", width=50)
        self.tree.column("Name", width=200)
        self.tree.column("Manager Code", width=100)
        self.tree.column("Type", width=100)
        self.tree.column("Start Date", width=100)
        self.tree.column("End Date", width=100)
        
        # Load initial data
        self.load_employees()
    
    def load_employees(self):
        """Load employees from database"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            session = get_session()
            employees = session.query(Employee).all()
            
            for emp in employees:
                self.tree.insert("", tk.END, values=(
                    emp.id,
                    emp.name,
                    emp.manager_code,
                    emp.employment_type,
                    emp.start_date,
                    emp.end_date
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employees: {str(e)}")
    
    def add_employee(self):
        """Open dialog to add a new employee"""
        dialog = EmployeeDialog(self)
        self.wait_window(dialog)
        
        if dialog.result:
            try:
                session = get_session()
                
                # Create new employee
                employee = Employee(
                    name=dialog.result["name"],
                    manager_code=dialog.result["manager_code"],
                    employment_type=dialog.result["employment_type"],
                    start_date=dialog.result["start_date"],
                    end_date=dialog.result["end_date"]
                )
                
                session.add(employee)
                session.commit()
                session.close()
                
                # Reload data
                self.load_employees()
                
                # Show success message
                messagebox.showinfo("Success", "Employee added successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add employee: {str(e)}")
    
    def edit_employee(self):
        """Open dialog to edit selected employee"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an employee to edit.")
            return
        
        # Get employee ID from selection
        emp_id = self.tree.item(selection[0])["values"][0]
        
        try:
            session = get_session()
            
            # Get the employee
            employee = session.query(Employee).filter(Employee.id == emp_id).first()
            if employee:
                # Open dialog with current values
                dialog = EmployeeDialog(self, employee)
                session.close()
                
                self.wait_window(dialog)
                
                if dialog.result:
                    try:
                        session = get_session()
                        
                        # Update employee
                        employee = session.query(Employee).filter(Employee.id == emp_id).first()
                        employee.name = dialog.result["name"]
                        employee.manager_code = dialog.result["manager_code"]
                        employee.employment_type = dialog.result["employment_type"]
                        employee.start_date = dialog.result["start_date"]
                        employee.end_date = dialog.result["end_date"]
                        
                        session.commit()
                        session.close()
                        
                        # Reload data
                        self.load_employees()
                        
                        # Show success message
                        messagebox.showinfo("Success", "Employee updated successfully.")
                        
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to update employee: {str(e)}")
            else:
                messagebox.showerror("Error", "Employee not found.")
                session.close()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employee: {str(e)}")
    
    def delete_employee(self):
        """Delete selected employee"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an employee to delete.")
            return
        
        # Get employee ID from selection
        emp_id = self.tree.item(selection[0])["values"][0]
        
        # Confirm deletion
        if not messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete employee with ID {emp_id}?"):
            return
        
        # Delete employee
        try:
            session = get_session()
            
            # Get the employee
            employee = session.query(Employee).filter(Employee.id == emp_id).first()
            if employee:
                session.delete(employee)
                session.commit()
                
                # Reload data
                self.load_employees()
                
                # Show success message
                messagebox.showinfo("Success", f"Employee with ID {emp_id} deleted successfully.")
            else:
                messagebox.showerror("Error", "Employee not found.")
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete employee: {str(e)}")

class ProjectAllocationTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="10")
        self.parent = parent
        self.create_widgets()
        
    def create_widgets(self):
        # Create toolbar
        toolbar = ttk.Frame(self, style="Panel.TFrame")
        toolbar.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Add decorative toolbar accent
        accent_canvas = tk.Canvas(toolbar, width=3, height=30, highlightthickness=0)
        accent_canvas.pack(side=tk.LEFT, padx=(0, 10), fill=tk.Y)
        # Draw a gradient accent
        for i in range(30):
            # Gradient from primary to secondary 
            r = int(74 + (52-74) * i/30)
            g = int(125 + (195-125) * i/30)
            b = int(186 + (143-186) * i/30)
            color = f'#{r:02x}{g:02x}{b:02x}'
            accent_canvas.create_line(0, i, 3, i, fill=color)
        
        # Button container for better visual grouping
        button_container = ttk.Frame(toolbar, style="Panel.TFrame")
        button_container.pack(side=tk.LEFT, fill=tk.Y)
        
        ttk.Button(button_container, text="Add Allocation", command=self.add_allocation).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Refresh", command=self.load_allocations).pack(side=tk.LEFT, padx=5)
        
        # Create treeview
        self.tree = ttk.Treeview(self, columns=("Manager", "Cost Center", "Work Code", "Year", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"), show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure treeview columns
        self.tree.heading("Manager", text="Manager")
        self.tree.heading("Cost Center", text="Cost Center")
        self.tree.heading("Work Code", text="Work Code")
        self.tree.heading("Year", text="Year")
        self.tree.heading("Jan", text="Jan")
        self.tree.heading("Feb", text="Feb")
        self.tree.heading("Mar", text="Mar")
        self.tree.heading("Apr", text="Apr")
        self.tree.heading("May", text="May")
        self.tree.heading("Jun", text="Jun")
        self.tree.heading("Jul", text="Jul")
        self.tree.heading("Aug", text="Aug")
        self.tree.heading("Sep", text="Sep")
        self.tree.heading("Oct", text="Oct")
        self.tree.heading("Nov", text="Nov")
        self.tree.heading("Dec", text="Dec")
        
        # Configure column widths
        self.tree.column("Manager", width=100)
        self.tree.column("Cost Center", width=100)
        self.tree.column("Work Code", width=100)
        self.tree.column("Year", width=60)
        self.tree.column("Jan", width=60)
        self.tree.column("Feb", width=60)
        self.tree.column("Mar", width=60)
        self.tree.column("Apr", width=60)
        self.tree.column("May", width=60)
        self.tree.column("Jun", width=60)
        self.tree.column("Jul", width=60)
        self.tree.column("Aug", width=60)
        self.tree.column("Sep", width=60)
        self.tree.column("Oct", width=60)
        self.tree.column("Nov", width=60)
        self.tree.column("Dec", width=60)
        
        # Load initial data
        self.load_allocations()
    
    def load_allocations(self):
        """Load project allocations from database"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            session = get_session()
            allocations = session.query(ProjectAllocation).all()
            
            for alloc in allocations:
                self.tree.insert("", tk.END, values=(
                    alloc.manager_code,
                    alloc.cost_center,
                    alloc.work_code_id,
                    alloc.year,
                    alloc.jan,
                    alloc.feb,
                    alloc.mar,
                    alloc.apr,
                    alloc.may,
                    alloc.jun,
                    alloc.jul,
                    alloc.aug,
                    alloc.sep,
                    alloc.oct,
                    alloc.nov,
                    alloc.dec
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load allocations: {str(e)}")
    
    def add_allocation(self):
        """Open dialog to add a new project allocation"""
        session = get_session()
        
        # Get the first manager code for demo
        # In a real app, you would have user select a manager
        manager = session.query(Employee.manager_code).first()
        manager_code = manager[0] if manager else ""
        session.close()
        
        if not manager_code:
            messagebox.showinfo("No Manager", "Please add employees with manager codes first.")
            return
            
        # Current year as default
        year = datetime.now().year
        
        # Open dialog
        dialog = ProjectAllocationDialog(self, manager_code, year)
        self.wait_window(dialog)
        
        if dialog.get_allocations():
            try:
                session = get_session()
                
                # Get allocations from dialog
                allocations = dialog.get_allocations()
                
                # Delete existing allocations for this manager and year
                session.query(ProjectAllocation).filter(
                    ProjectAllocation.manager_code == manager_code,
                    ProjectAllocation.year == year
                ).delete()
                
                # Add new allocations
                for allocation in allocations:
                    project_allocation = ProjectAllocation(
                        manager_code=manager_code,
                        cost_center=allocation['cost_center'],
                        work_code_id=allocation['work_code_id'],
                        year=year,
                        jan=allocation['jan'],
                        feb=allocation['feb'],
                        mar=allocation['mar'],
                        apr=allocation['apr'],
                        may=allocation['may'],
                        jun=allocation['jun'],
                        jul=allocation['jul'],
                        aug=allocation['aug'],
                        sep=allocation['sep'],
                        oct=allocation['oct'],
                        nov=allocation['nov'],
                        dec=allocation['dec']
                    )
                    session.add(project_allocation)
                
                session.commit()
                session.close()
                
                # Refresh view
                self.load_allocations()
                
                # Show success message
                messagebox.showinfo("Success", "Project allocations updated successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update allocations: {str(e)}")

class ProjectAllocationDialog(tk.Toplevel):
    def __init__(self, parent, manager_code, year):
        super().__init__(parent)
        self.manager_code = manager_code
        self.year = year
        self.result = None
        
        self.title("Project Allocation")
        self.geometry("800x600")
        self.resizable(True, True)
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Manager and Year info
        info_frame = ttk.LabelFrame(frame, text="Allocation Info", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(info_frame, text=f"Manager Code: {manager_code}").pack(side=tk.LEFT, padx=5)
        ttk.Label(info_frame, text=f"Year: {year}").pack(side=tk.LEFT, padx=5)
        
        # Project details
        project_frame = ttk.LabelFrame(frame, text="Project Details", padding="10")
        project_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(project_frame, text="Cost Center:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(project_frame, textvariable=self.cost_center_var, width=20).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(project_frame, text="Work Code:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.work_code_var = tk.StringVar()
        ttk.Entry(project_frame, textvariable=self.work_code_var, width=20).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Monthly allocations
        allocation_frame = ttk.LabelFrame(frame, text="Monthly Allocations", padding="10")
        allocation_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create month spinboxes
        self.spinboxes = {}
        months = [
            ("January", "jan"), ("February", "feb"), ("March", "mar"),
            ("April", "apr"), ("May", "may"), ("June", "jun"),
            ("July", "jul"), ("August", "aug"), ("September", "sep"),
            ("October", "oct"), ("November", "nov"), ("December", "dec")
        ]
        
        for i, (month_name, month_code) in enumerate(months):
            row = i // 3
            col = i % 3
            
            month_frame = ttk.Frame(allocation_frame)
            month_frame.grid(row=row, column=col, padx=5, pady=5, sticky=tk.W)
            
            ttk.Label(month_frame, text=f"{month_name}:").pack(side=tk.LEFT)
            spinbox = ttk.Spinbox(month_frame, from_=0, to=100, increment=0.5, width=10)
            spinbox.pack(side=tk.LEFT, padx=5)
            self.spinboxes[month_code] = spinbox
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Load current allocations
        self.load_allocations()
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)
    
    def load_allocations(self):
        """Load project allocations from database"""
        session = get_session()
        allocations = session.query(ProjectAllocation).filter(
            ProjectAllocation.manager_code == self.manager_code,
            ProjectAllocation.year == self.year
        ).all()
        
        for allocation in allocations:
            if allocation.cost_center:
                self.cost_center_var.set(allocation.cost_center)
            if allocation.work_code_id:
                self.work_code_var.set(allocation.work_code_id)
            
            # Set monthly values
            self.spinboxes["jan"].set(allocation.jan or 0)
            self.spinboxes["feb"].set(allocation.feb or 0)
            self.spinboxes["mar"].set(allocation.mar or 0)
            self.spinboxes["apr"].set(allocation.apr or 0)
            self.spinboxes["may"].set(allocation.may or 0)
            self.spinboxes["jun"].set(allocation.jun or 0)
            self.spinboxes["jul"].set(allocation.jul or 0)
            self.spinboxes["aug"].set(allocation.aug or 0)
            self.spinboxes["sep"].set(allocation.sep or 0)
            self.spinboxes["oct"].set(allocation.oct or 0)
            self.spinboxes["nov"].set(allocation.nov or 0)
            self.spinboxes["dec"].set(allocation.dec or 0)
        
        session.close()
    
    def on_ok(self):
        try:
            # Validate required fields
            cost_center = self.cost_center_var.get().strip()
            work_code = self.work_code_var.get().strip()
            
            if not cost_center:
                raise ValueError("Cost Center is required")
            if not work_code:
                raise ValueError("Work Code is required")
            
            # Get monthly values
            monthly_values = {}
            for month, spinbox in self.spinboxes.items():
                try:
                    monthly_values[month] = float(spinbox.get())
                except ValueError:
                    raise ValueError(f"Invalid value for {month}")
            
            self.result = {
                "cost_center": cost_center,
                "work_code_id": work_code,
                **monthly_values
            }
            self.destroy()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
    
    def on_cancel(self):
        self.result = None
        self.destroy()
    
    def get_allocations(self):
        return self.result

class ForecastVisualization(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.create_widgets()
        
    def create_widgets(self):
        # Control frame for options
        control_frame = ttk.Frame(self)
        control_frame.pack(fill=tk.X, expand=False, pady=5)
        
        # Year selector
        ttk.Label(control_frame, text="Year:").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        years = [str(year) for year in range(datetime.now().year-2, datetime.now().year+3)]
        self.year_combo = ttk.Combobox(control_frame, textvariable=self.year_var, values=years, width=6)
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
        # Chart type selector
        ttk.Label(control_frame, text="Chart Type:").pack(side=tk.LEFT, padx=5)
        self.chart_type_var = tk.StringVar(value="Monthly Forecast")
        chart_types = ["Monthly Forecast", "Manager Allocation", "Employee Type Distribution"]
        self.chart_type_combo = ttk.Combobox(control_frame, textvariable=self.chart_type_var, 
                                        values=chart_types, width=20, state="readonly")
        self.chart_type_combo.pack(side=tk.LEFT, padx=5)
        
        # Create button
        ttk.Button(control_frame, text="Generate Chart", command=self.generate_chart).pack(side=tk.LEFT, padx=10)
        
        # Frame for the chart
        self.chart_frame = ttk.Frame(self)
        self.chart_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Initial message
        ttk.Label(self.chart_frame, text="Select options and click 'Generate Chart' to create a visualization").pack(
            pady=50)
    
    def generate_chart(self):
        # Clear previous chart
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        
        try:
            chart_type = self.chart_type_var.get()
            year = int(self.year_var.get())
            
            # Create figure and axes
            fig = Figure(figsize=(10, 6), dpi=100)
            ax = fig.add_subplot(111)
            
            if chart_type == "Monthly Forecast":
                self._generate_monthly_forecast_chart(ax, year)
            elif chart_type == "Manager Allocation":
                self._generate_manager_allocation_chart(ax, year)
            elif chart_type == "Employee Type Distribution":
                self._generate_employee_type_distribution(ax, year)
            
            # Add the plot to the window
            canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate chart: {str(e)}")
            # Re-add the initial message
            ttk.Label(self.chart_frame, text=f"Error generating chart: {str(e)}").pack(pady=50)
    
    def _generate_monthly_forecast_chart(self, ax, year):
        session = get_session()
        
        # Get forecast data for the year
        forecasts = session.query(Forecast).filter(Forecast.year == year).all()
        
        # Group by month
        monthly_data = {}
        for month in range(1, 13):
            monthly_data[month] = 0
            
        for forecast in forecasts:
            monthly_data[forecast.month] += forecast.hours
        
        # Create the chart
        months = list(monthly_data.keys())
        hours = list(monthly_data.values())
        
        ax.bar(months, hours, color='skyblue')
        ax.set_xlabel('Month')
        ax.set_ylabel('Hours')
        ax.set_title(f'Monthly Forecast Hours for {year}')
        ax.set_xticks(months)
        ax.set_xticklabels(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'])
        ax.grid(axis='y', linestyle='--', alpha=0.7)
        
        session.close()
    
    def _generate_manager_allocation_chart(self, ax, year):
        session = get_session()
        
        # Get project allocations for the year
        allocations = session.query(ProjectAllocation).filter(ProjectAllocation.year == year).all()
        
        # Group by manager code
        manager_data = {}
        for allocation in allocations:
            if allocation.manager_code not in manager_data:
                manager_data[allocation.manager_code] = 0
            manager_data[allocation.manager_code] += allocation.hours
        
        # Sort by hours (descending)
        sorted_data = dict(sorted(manager_data.items(), key=lambda item: item[1], reverse=True))
        
        managers = list(sorted_data.keys())
        hours = list(sorted_data.values())
        
        # Limit to top 10 if there are many
        if len(managers) > 10:
            managers = managers[:10]
            hours = hours[:10]
            ax.set_title(f'Top 10 Manager Allocations for {year}')
        else:
            ax.set_title(f'Manager Allocations for {year}')
        
        # Create the chart
        ax.barh(managers, hours, color='lightgreen')
        ax.set_xlabel('Hours')
        ax.set_ylabel('Manager Code')
        ax.grid(axis='x', linestyle='--', alpha=0.7)
        
        session.close()
    
    def _generate_employee_type_distribution(self, ax, year):
        session = get_session()
        
        # Get all employees
        employees = session.query(Employee).all()
        
        # Count by type
        fte_count = sum(1 for e in employees if e.employment_type == "FTE")
        contractor_count = sum(1 for e in employees if e.employment_type == "Contractor")
        
        # Create pie chart
        labels = ['FTE', 'Contractor']
        sizes = [fte_count, contractor_count]
        colors = ['#ff9999','#66b3ff']
        explode = (0.1, 0)  # explode the 1st slice (FTE)
        
        ax.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', shadow=True, startangle=90)
        ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
        ax.set_title(f'Employee Type Distribution')
        
        session.close()

class GA01WeeksTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="10")
        self.parent = parent
        self.create_widgets()
        
    def create_widgets(self):
        # Create toolbar
        toolbar = ttk.Frame(self, style="Panel.TFrame")
        toolbar.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Add decorative toolbar accent
        accent_canvas = tk.Canvas(toolbar, width=3, height=30, highlightthickness=0)
        accent_canvas.pack(side=tk.LEFT, padx=(0, 10), fill=tk.Y)
        # Draw a gradient accent
        for i in range(30):
            r = int(74 + (52-74) * i/30)
            g = int(125 + (195-125) * i/30)
            b = int(186 + (143-186) * i/30)
            color = f'#{r:02x}{g:02x}{b:02x}'
            accent_canvas.create_line(0, i, 3, i, fill=color)
        
        # Button container for better visual grouping
        button_container = ttk.Frame(toolbar, style="Panel.TFrame")
        button_container.pack(side=tk.LEFT, fill=tk.Y)
        
        # Year selector
        ttk.Label(button_container, text="Year:").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        years = [str(year) for year in range(datetime.now().year-2, datetime.now().year+3)]
        self.year_combo = ttk.Combobox(button_container, textvariable=self.year_var, values=years, width=6)
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_container, text="Load", command=self.load_ga01_weeks).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Save", command=self.save_ga01_weeks).pack(side=tk.LEFT, padx=5)
        
        # Create main frame for month entries
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Month entries
        self.spinboxes = {}
        
        for i, month in enumerate(range(1, 13)):
            month_name = datetime(2000, month, 1).strftime('%B')
            ttk.Label(main_frame, text=f"{month_name}:").grid(row=i, column=0, sticky=tk.W, pady=5, padx=5)
            
            var = tk.DoubleVar(value=4.0)
            spinbox = ttk.Spinbox(main_frame, from_=0, to=5, increment=0.01, textvariable=var, width=10)
            spinbox.grid(row=i, column=1, sticky=tk.W, pady=5, padx=5)
            self.spinboxes[month] = var
        
        # Load initial data
        self.load_ga01_weeks()
    
    def load_ga01_weeks(self):
        """Load GA01 weeks from database"""
        try:
            year = int(self.year_var.get())
            session = get_session()
            ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
            
            # Reset to default values
            for month in range(1, 13):
                self.spinboxes[month].set(4.0)
            
            # Update with database values
            for week in ga01_weeks:
                if week.month in self.spinboxes:
                    self.spinboxes[week.month].set(week.weeks)
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load GA01 weeks: {str(e)}")
    
    def save_ga01_weeks(self):
        """Save GA01 weeks to database"""
        try:
            year = int(self.year_var.get())
            session = get_session()
            
            # Delete existing weeks for this year
            session.query(GA01Week).filter(GA01Week.year == year).delete()
            
            # Add new weeks
            for month, var in self.spinboxes.items():
                ga01_week = GA01Week(
                    year=year,
                    month=month,
                    weeks=var.get()
                )
                session.add(ga01_week)
            
            session.commit()
            session.close()
            
            messagebox.showinfo("Success", f"GA01 weeks for {year} saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save GA01 weeks: {str(e)}")

class PlannedChangesTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="10")
        self.parent = parent
        self.create_widgets()
        
    def create_widgets(self):
        # Create toolbar
        toolbar = ttk.Frame(self, style="Panel.TFrame")
        toolbar.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Add decorative toolbar accent
        accent_canvas = tk.Canvas(toolbar, width=3, height=30, highlightthickness=0)
        accent_canvas.pack(side=tk.LEFT, padx=(0, 10), fill=tk.Y)
        # Draw a gradient accent
        for i in range(30):
            r = int(74 + (52-74) * i/30)
            g = int(125 + (195-125) * i/30)
            b = int(186 + (143-186) * i/30)
            color = f'#{r:02x}{g:02x}{b:02x}'
            accent_canvas.create_line(0, i, 3, i, fill=color)
        
        # Button container for better visual grouping
        button_container = ttk.Frame(toolbar, style="Panel.TFrame")
        button_container.pack(side=tk.LEFT, fill=tk.Y)
        
        ttk.Button(button_container, text="Add Change", command=self.add_change).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Edit Change", command=self.edit_change).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Delete Change", command=self.delete_change).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Refresh", command=self.load_changes).pack(side=tk.LEFT, padx=5)
        
        # Create treeview
        self.tree = ttk.Treeview(self, columns=("ID", "Description", "Type", "Effective Date", "Status"), show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure treeview columns
        self.tree.heading("ID", text="ID")
        self.tree.heading("Description", text="Description")
        self.tree.heading("Type", text="Type")
        self.tree.heading("Effective Date", text="Effective Date")
        self.tree.heading("Status", text="Status")
        
        # Configure column widths
        self.tree.column("ID", width=50)
        self.tree.column("Description", width=300)
        self.tree.column("Type", width=100)
        self.tree.column("Effective Date", width=100)
        self.tree.column("Status", width=100)
        
        # Load initial data
        self.load_changes()
    
    def load_changes(self):
        """Load planned changes from database"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            session = get_session()
            changes = session.query(PlannedChange).all()
            
            for change in changes:
                self.tree.insert("", tk.END, values=(
                    change.id,
                    change.description,
                    change.change_type,
                    change.effective_date,
                    change.status
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load planned changes: {str(e)}")
    
    def add_change(self):
        """Open dialog to add a new planned change"""
        dialog = PlannedChangeDialog(self)
        self.wait_window(dialog)
        
        if dialog.result:
            try:
                session = get_session()
                
                # Create new planned change
                change = PlannedChange(
                    description=dialog.result["description"],
                    change_type=dialog.result["change_type"],
                    effective_date=dialog.result["effective_date"],
                    employee_id=dialog.result.get("employee_id"),
                    target_type=dialog.result.get("target_type"),
                    name=dialog.result.get("name"),
                    team=dialog.result.get("team"),
                    manager_code=dialog.result.get("manager_code"),
                    cost_center=dialog.result.get("cost_center"),
                    employment_type=dialog.result.get("employment_type"),
                    status="Pending"
                )
                
                session.add(change)
                session.commit()
                session.close()
                
                # Reload data
                self.load_changes()
                
                # Show success message
                messagebox.showinfo("Success", "Planned change added successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add planned change: {str(e)}")
    
    def edit_change(self):
        """Open dialog to edit selected planned change"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a planned change to edit.")
            return
        
        # Get change ID from selection
        change_id = self.tree.item(selection[0])["values"][0]
        
        try:
            session = get_session()
            
            # Get the planned change
            change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
            if change:
                # Open dialog with current values
                dialog = PlannedChangeDialog(self, change)
                session.close()
                
                self.wait_window(dialog)
                
                if dialog.result:
                    try:
                        session = get_session()
                        
                        # Update planned change
                        change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
                        change.description = dialog.result["description"]
                        change.change_type = dialog.result["change_type"]
                        change.effective_date = dialog.result["effective_date"]
                        change.employee_id = dialog.result.get("employee_id")
                        change.target_type = dialog.result.get("target_type")
                        change.name = dialog.result.get("name")
                        change.team = dialog.result.get("team")
                        change.manager_code = dialog.result.get("manager_code")
                        change.cost_center = dialog.result.get("cost_center")
                        change.employment_type = dialog.result.get("employment_type")
                        
                        session.commit()
                        session.close()
                        
                        # Reload data
                        self.load_changes()
                        
                        # Show success message
                        messagebox.showinfo("Success", "Planned change updated successfully.")
                        
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to update planned change: {str(e)}")
            else:
                messagebox.showerror("Error", "Planned change not found.")
                session.close()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load planned change: {str(e)}")
    
    def delete_change(self):
        """Delete selected planned change"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a planned change to delete.")
            return
        
        # Get change ID from selection
        change_id = self.tree.item(selection[0])["values"][0]
        
        # Confirm deletion
        if not messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete planned change with ID {change_id}?"):
            return
        
        try:
            session = get_session()
            
            # Get the planned change
            change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
            if change:
                session.delete(change)
                session.commit()
                
                # Reload data
                self.load_changes()
                
                # Show success message
                messagebox.showinfo("Success", f"Planned change with ID {change_id} deleted successfully.")
            else:
                messagebox.showerror("Error", "Planned change not found.")
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete planned change: {str(e)}")

class ForecastTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="10")
        self.parent = parent
        self.create_widgets()
        
    def create_widgets(self):
        # Create toolbar
        toolbar = ttk.Frame(self, style="Panel.TFrame")
        toolbar.pack(fill=tk.X, pady=(0, 10), padx=5)
        
        # Add decorative toolbar accent
        accent_canvas = tk.Canvas(toolbar, width=3, height=30, highlightthickness=0)
        accent_canvas.pack(side=tk.LEFT, padx=(0, 10), fill=tk.Y)
        # Draw a gradient accent
        for i in range(30):
            r = int(74 + (52-74) * i/30)
            g = int(125 + (195-125) * i/30)
            b = int(186 + (143-186) * i/30)
            color = f'#{r:02x}{g:02x}{b:02x}'
            accent_canvas.create_line(0, i, 3, i, fill=color)
        
        # Button container for better visual grouping
        button_container = ttk.Frame(toolbar, style="Panel.TFrame")
        button_container.pack(side=tk.LEFT, fill=tk.Y)
        
        # Year selector
        ttk.Label(button_container, text="Year:").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        years = [str(year) for year in range(datetime.now().year-2, datetime.now().year+3)]
        self.year_combo = ttk.Combobox(button_container, textvariable=self.year_var, values=years, width=6)
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
        # Work code selector
        ttk.Label(button_container, text="Work Code:").pack(side=tk.LEFT, padx=5)
        self.work_code_var = tk.StringVar(value="All")
        self.work_code_combo = ttk.Combobox(button_container, textvariable=self.work_code_var, width=20)
        self.work_code_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_container, text="Calculate Forecast", command=self.calculate_forecast).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Refresh", command=self.load_forecast).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_container, text="Export to Excel", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)
        
        # Create treeview
        self.tree = ttk.Treeview(self, columns=("Cost Center", "Manager", "Work Code", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total"), show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure treeview columns
        self.tree.heading("Cost Center", text="Cost Center")
        self.tree.heading("Manager", text="Manager")
        self.tree.heading("Work Code", text="Work Code")
        self.tree.heading("Jan", text="Jan")
        self.tree.heading("Feb", text="Feb")
        self.tree.heading("Mar", text="Mar")
        self.tree.heading("Apr", text="Apr")
        self.tree.heading("May", text="May")
        self.tree.heading("Jun", text="Jun")
        self.tree.heading("Jul", text="Jul")
        self.tree.heading("Aug", text="Aug")
        self.tree.heading("Sep", text="Sep")
        self.tree.heading("Oct", text="Oct")
        self.tree.heading("Nov", text="Nov")
        self.tree.heading("Dec", text="Dec")
        self.tree.heading("Total", text="Total")
        
        # Configure column widths
        self.tree.column("Cost Center", width=100)
        self.tree.column("Manager", width=100)
        self.tree.column("Work Code", width=100)
        self.tree.column("Jan", width=60)
        self.tree.column("Feb", width=60)
        self.tree.column("Mar", width=60)
        self.tree.column("Apr", width=60)
        self.tree.column("May", width=60)
        self.tree.column("Jun", width=60)
        self.tree.column("Jul", width=60)
        self.tree.column("Aug", width=60)
        self.tree.column("Sep", width=60)
        self.tree.column("Oct", width=60)
        self.tree.column("Nov", width=60)
        self.tree.column("Dec", width=60)
        self.tree.column("Total", width=80)
        
        # Load initial data
        self.load_work_codes()
        self.load_forecast()
    
    def load_work_codes(self):
        """Load work codes from project allocations"""
        try:
            session = get_session()
            work_codes = session.query(ProjectAllocation.work_code_id).distinct().all()
            work_codes = ["All"] + [wc[0] for wc in work_codes]
            self.work_code_combo["values"] = work_codes
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load work codes: {str(e)}")
    
    def calculate_forecast(self):
        """Calculate forecast based on project allocations and GA01 weeks"""
        try:
            year = int(self.year_var.get())
            work_code = self.work_code_var.get()
            
            session = get_session()
            
            # Get GA01 weeks for the year
            ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
            ga01_dict = {week.month: week.weeks for week in ga01_weeks}
            
            # Get project allocations
            query = session.query(ProjectAllocation).filter(ProjectAllocation.year == year)
            if work_code != "All":
                query = query.filter(ProjectAllocation.work_code_id == work_code)
            allocations = query.all()
            
            # Calculate forecast
            forecasts = []
            for allocation in allocations:
                forecast = Forecast(
                    year=year,
                    cost_center=allocation.cost_center,
                    manager_code=allocation.manager_code,
                    work_code_id=allocation.work_code_id
                )
                
                # Calculate monthly hours based on GA01 weeks
                total_hours = 0
                for month in range(1, 13):
                    weeks = ga01_dict.get(month, 4.0)  # Default to 4 weeks if not set
                    hours = getattr(allocation, datetime(2000, month, 1).strftime('%b').lower()) * weeks
                    setattr(forecast, datetime(2000, month, 1).strftime('%b').lower(), hours)
                    total_hours += hours
                
                forecast.total_hours = total_hours
                forecasts.append(forecast)
            
            # Save forecasts
            session.query(Forecast).filter(Forecast.year == year).delete()
            for forecast in forecasts:
                session.add(forecast)
            
            session.commit()
            session.close()
            
            # Reload data
            self.load_forecast()
            
            # Show success message
            messagebox.showinfo("Success", "Forecast calculated successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate forecast: {str(e)}")
    
    def load_forecast(self):
        """Load forecast data from database"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            year = int(self.year_var.get())
            work_code = self.work_code_var.get()
            
            session = get_session()
            
            # Get forecasts
            query = session.query(Forecast).filter(Forecast.year == year)
            if work_code != "All":
                query = query.filter(Forecast.work_code_id == work_code)
            forecasts = query.all()
            
            for forecast in forecasts:
                self.tree.insert("", tk.END, values=(
                    forecast.cost_center,
                    forecast.manager_code,
                    forecast.work_code_id,
                    forecast.jan,
                    forecast.feb,
                    forecast.mar,
                    forecast.apr,
                    forecast.may,
                    forecast.jun,
                    forecast.jul,
                    forecast.aug,
                    forecast.sep,
                    forecast.oct,
                    forecast.nov,
                    forecast.dec,
                    forecast.total_hours
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load forecast: {str(e)}")
    
    def export_to_excel(self):
        """Export forecast data to Excel"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill
            
            # Create workbook and select active sheet
            wb = Workbook()
            ws = wb.active
            ws.title = "Forecast"
            
            # Write headers
            headers = ["Cost Center", "Manager", "Work Code", "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col)
                cell.value = header
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Write data
            row = 2
            for item in self.tree.get_children():
                values = self.tree.item(item)["values"]
                for col, value in enumerate(values, 1):
                    ws.cell(row=row, column=col, value=value)
                row += 1
            
            # Save file
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Export Forecast"
            )
            
            if filename:
                wb.save(filename)
                messagebox.showinfo("Success", f"Forecast exported to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export forecast: {str(e)}")

class ForecastApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Forecast Tool")
        self.geometry("1200x800")
        
        # Configure modern style
        style = ttk.Style()
        style.configure(".", font=("Helvetica", 10))
        style.configure("Treeview", rowheight=25)
        style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))
        style.configure("TButton", padding=6)
        style.configure("TEntry", padding=6)
        style.configure("TCombobox", padding=6)
        style.configure("Panel.TFrame", background="#f0f0f0")
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.employee_tab = EmployeeTab(self.notebook)
        self.notebook.add(self.employee_tab, text="Employees")
        
        self.ga01_weeks_tab = GA01WeeksTab(self.notebook)
        self.notebook.add(self.ga01_weeks_tab, text="GA01 Weeks")
        
        self.forecast_tab = ForecastTab(self.notebook)
        self.notebook.add(self.forecast_tab, text="Forecast")
        
        self.project_allocation_tab = ProjectAllocationTab(self.notebook)
        self.notebook.add(self.project_allocation_tab, text="Project Allocations")
        
        self.planned_changes_tab = PlannedChangesTab(self.notebook)
        self.notebook.add(self.planned_changes_tab, text="Planned Changes")
        
        self.forecast_visualization_tab = ForecastVisualization(self.notebook)
        self.notebook.add(self.forecast_visualization_tab, text="Visualizations")
        
        # Create menu
        self.create_menu()
        
        # Create status bar
        self.status_bar = ttk.Label(self, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_menu(self):
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Employees", command=self.import_employees)
        file_menu.add_command(label="Import GA01 Weeks", command=self.import_ga01_weeks)
        file_menu.add_command(label="Export Forecast", command=self.export_forecast)
        file_menu.add_separator()
        file_menu.add_command(label="Settings", command=self.open_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def import_employees(self):
        messagebox.showinfo("Not Implemented", "Import Employees feature coming soon!")
    
    def import_ga01_weeks(self):
        messagebox.showinfo("Not Implemented", "Import GA01 Weeks feature coming soon!")
    
    def export_forecast(self):
        messagebox.showinfo("Not Implemented", "Export Forecast feature coming soon!")
    
    def open_settings(self):
        dialog = SettingsDialog(self)
        if dialog.get_settings():
            try:
                session = get_session()
                
                # Get or create settings
                settings = session.query(Settings).first()
                if not settings:
                    settings = Settings()
                    session.add(settings)
                
                # Update settings
                settings.fte_hours = dialog.get_settings()["fte_hours"]
                settings.contractor_hours = dialog.get_settings()["contractor_hours"]
                
                session.commit()
                session.close()
                
                messagebox.showinfo("Success", "Settings updated successfully.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update settings: {str(e)}")
    
    def show_about(self):
        messagebox.showinfo("About", "Forecast Tool v1.0\n\nA tool for managing employee forecasts.")

if __name__ == "__main__":
    app = ForecastApp()
    app.mainloop() 