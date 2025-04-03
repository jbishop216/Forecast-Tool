import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, Enum
from sqlalchemy.orm import declarative_base, sessionmaker
import enum
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import csv
import os

# Define modern color scheme
COLORS = {
    'primary': '#2c3e50',      # Dark blue-gray
    'secondary': '#3498db',    # Bright blue
    'accent': '#e74c3c',       # Red
    'success': '#2ecc71',      # Green
    'warning': '#f1c40f',      # Yellow
    'background': '#ecf0f1',   # Light gray
    'text': '#000000',         # Black (changed from blue-gray for better readability)
    'text_light': '#34495e',   # Darker gray (changed for better contrast)
    'white': '#ffffff',        # White
    'border': '#bdc3c7'        # Medium gray
}

# Configure ttk styles
def configure_styles():
    style = ttk.Style()
    
    # Configure main window style
    style.configure('Main.TFrame', background=COLORS['background'])
    
    # Configure notebook style
    style.configure('TNotebook', background=COLORS['background'])
    style.configure('TNotebook.Tab', padding=[10, 5], background=COLORS['background'])
    style.map('TNotebook.Tab',
        background=[('selected', COLORS['primary'])],
        foreground=[('selected', COLORS['white'])]
    )
    
    # Configure button styles
    style.configure('Primary.TButton',
        padding=10,
        background=COLORS['primary'],
        foreground=COLORS['white']
    )
    style.map('Primary.TButton',
        background=[('active', COLORS['secondary'])],
        foreground=[('active', COLORS['white'])]
    )
    
    # Configure entry styles
    style.configure('TEntry',
        padding=5,
        fieldbackground=COLORS['white'],
        borderwidth=1
    )
    
    # Configure label styles
    style.configure('TLabel',
        background=COLORS['background'],
        foreground=COLORS['text'],
        padding=5
    )
    
    # Configure treeview styles
    style.configure('Treeview',
        background=COLORS['white'],
        foreground=COLORS['text'],
        fieldbackground=COLORS['white'],
        rowheight=25
    )
    style.configure('Treeview.Heading',
        background=COLORS['primary'],
        foreground=COLORS['white'],
        padding=5
    )
    style.map('Treeview',
        background=[('selected', COLORS['secondary'])],
        foreground=[('selected', COLORS['white'])]
    )
    
    # Configure combobox styles
    style.configure('TCombobox',
        background=COLORS['white'],
        fieldbackground=COLORS['white'],
        selectbackground=COLORS['secondary'],
        selectforeground=COLORS['white']
    )
    
    # Configure frame styles
    style.configure('Card.TFrame',
        background=COLORS['white'],
        relief='solid',
        borderwidth=1
    )
    
    # Configure separator style
    style.configure('TSeparator',
        background=COLORS['border']
    )

# Create database engine
engine = create_engine('sqlite:///forecast_tool.db')
Base = declarative_base()
Session = sessionmaker(bind=engine)

def get_session():
    return Session()

def verify_db_connection():
    """Verify database connection and reset it if necessary"""
    try:
        session = get_session()
        # Try a simple query
        session.query(Settings).first()
        session.close()
        return True
    except Exception as e:
        print(f"Database connection error: {str(e)}")
        try:
            # Try to reset the connection
            global engine, Session
            engine = create_engine('sqlite:///forecast_tool.db')
            Session = sessionmaker(bind=engine)
            # Create tables if they don't exist
            Base.metadata.create_all(engine)
            return True
        except Exception as e:
            print(f"Failed to reset database connection: {str(e)}")
            return False

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
    cost_center = Column(String, nullable=False)
    employment_type = Column(String, nullable=False)
    start_date = Column(Date, nullable=False)
    end_date = Column(Date)

class ProjectAllocation(Base):
    __tablename__ = 'project_allocations'
    
    id = Column(Integer, primary_key=True)
    manager_code = Column(String, nullable=False)
    cost_center = Column(String, nullable=False)
    work_code = Column(String, nullable=False)
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
    work_code = Column(String, nullable=False)
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
        self.parent = parent
        self.employee = employee
        self.result = None
        
        self.title("Employee")
        self.geometry("450x400")
        self.resizable(True, True)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
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
        
        # Cost Center
        ttk.Label(frame, text="Cost Center:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.cost_center_var, width=20).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Employment Type
        ttk.Label(frame, text="Employment Type:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.employment_type_var = tk.StringVar()
        employment_types = [et.value for et in EmploymentType]
        ttk.Combobox(frame, textvariable=self.employment_type_var, values=employment_types, width=20, state="readonly").grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Start Date
        ttk.Label(frame, text="Start Date (mm/dd/yy):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.start_date_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.start_date_var, width=20).grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # End Date
        ttk.Label(frame, text="End Date (mm/dd/yy):").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.end_date_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.end_date_var, width=20).grid(row=5, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # If editing, populate fields
        if employee:
            self.name_var.set(employee.name)
            self.manager_code_var.set(employee.manager_code)
            self.cost_center_var.set(employee.cost_center)
            self.employment_type_var.set(employee.employment_type)
            self.start_date_var.set(employee.start_date.strftime("%m/%d/%y") if employee.start_date else "")
            self.end_date_var.set(employee.end_date.strftime("%m/%d/%y") if employee.end_date else "")
        else:
            self.start_date_var.set(datetime.now().strftime("%m/%d/%y"))
        
        # Center the dialog on the parent window
        self.center_on_parent()
    
    def center_on_parent(self):
        """Center the dialog on its parent window"""
        self.update_idletasks()
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        width = self.winfo_width()
        height = self.winfo_height()
        
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        self.geometry(f"+{x}+{y}")
    
    def on_ok(self):
        try:
            # Validate required fields
            if not self.name_var.get().strip():
                raise ValueError("Name is required")
            if not self.manager_code_var.get().strip():
                raise ValueError("Manager Code is required")
            if not self.cost_center_var.get().strip():
                raise ValueError("Cost Center is required")
            if not self.employment_type_var.get():
                raise ValueError("Employment Type is required")
            if not self.start_date_var.get().strip():
                raise ValueError("Start Date is required")
            
            # Parse dates
            start_date = datetime.strptime(self.start_date_var.get(), "%m/%d/%y").date()
            end_date = None
            if self.end_date_var.get().strip():
                end_date = datetime.strptime(self.end_date_var.get(), "%m/%d/%y").date()
                if end_date <= start_date:
                    raise ValueError("End Date must be after Start Date")
            
            self.result = {
                "name": self.name_var.get().strip(),
                "manager_code": self.manager_code_var.get().strip(),
                "cost_center": self.cost_center_var.get().strip(),
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
        self.tree = ttk.Treeview(self, columns=("ID", "Name", "Manager Code", "Cost Center", "Type", "Start Date", "End Date"), show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configure treeview columns
        self.tree.heading("ID", text="ID")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Manager Code", text="Manager Code")
        self.tree.heading("Cost Center", text="Cost Center")
        self.tree.heading("Type", text="Type")
        self.tree.heading("Start Date", text="Start Date")
        self.tree.heading("End Date", text="End Date")
        
        # Configure column widths
        self.tree.column("ID", width=50)
        self.tree.column("Name", width=200)
        self.tree.column("Manager Code", width=100)
        self.tree.column("Cost Center", width=100)
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
                    emp.cost_center,
                    emp.employment_type,
                    emp.start_date.strftime("%m/%d/%y") if emp.start_date else "",
                    emp.end_date.strftime("%m/%d/%y") if emp.end_date else ""
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employees: {str(e)}")
    
    def add_employee(self):
        """Open dialog to add a new employee"""
        dialog = EmployeeDialog(self)
        dialog.wait_window()
        
        if dialog.result:
            try:
                session = get_session()
                
                # Create new employee
                employee = Employee(
                    name=dialog.result["name"],
                    manager_code=dialog.result["manager_code"],
                    cost_center=dialog.result["cost_center"],
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
                
                dialog.wait_window()
                
                if dialog.result:
                    try:
                        session = get_session()
                        
                        # Update employee
                        employee = session.query(Employee).filter(Employee.id == emp_id).first()
                        employee.name = dialog.result["name"]
                        employee.manager_code = dialog.result["manager_code"]
                        employee.cost_center = dialog.result["cost_center"]
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
        super().__init__(parent)
        
        # Create toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Add Allocation", command=self.add_allocation).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Allocation", command=self.edit_allocation).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Allocation", command=self.delete_allocation).pack(side=tk.LEFT, padx=2)
        
        # Create treeview with scrollbar
        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=(
            "manager_code", "year", "cost_center", "work_code",
            "jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "oct", "nov", "dec"
        ), show="headings")
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        self.tree.heading("manager_code", text="Manager")
        self.tree.heading("year", text="Year")
        self.tree.heading("cost_center", text="Cost Center")
        self.tree.heading("work_code", text="Work Code")
        self.tree.heading("jan", text="Jan")
        self.tree.heading("feb", text="Feb")
        self.tree.heading("mar", text="Mar")
        self.tree.heading("apr", text="Apr")
        self.tree.heading("may", text="May")
        self.tree.heading("jun", text="Jun")
        self.tree.heading("jul", text="Jul")
        self.tree.heading("aug", text="Aug")
        self.tree.heading("sep", text="Sep")
        self.tree.heading("oct", text="Oct")
        self.tree.heading("nov", text="Nov")
        self.tree.heading("dec", text="Dec")
        
        # Set column widths
        self.tree.column("manager_code", width=100)
        self.tree.column("year", width=60)
        self.tree.column("cost_center", width=100)
        self.tree.column("work_code", width=100)
        for month in ["jan", "feb", "mar", "apr", "may", "jun",
                     "jul", "aug", "sep", "oct", "nov", "dec"]:
            self.tree.column(month, width=50)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Load allocations
        self.load_allocations()
    
    def load_allocations(self):
        """Load project allocations from database"""
        try:
            # Clear existing items
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            session = get_session()
            allocations = session.query(ProjectAllocation).all()
            
            for allocation in allocations:
                self.tree.insert("", tk.END, values=(
                    allocation.manager_code,
                    allocation.year,
                    allocation.cost_center,
                    allocation.work_code,
                    allocation.jan,
                    allocation.feb,
                    allocation.mar,
                    allocation.apr,
                    allocation.may,
                    allocation.jun,
                    allocation.jul,
                    allocation.aug,
                    allocation.sep,
                    allocation.oct,
                    allocation.nov,
                    allocation.dec
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load allocations: {str(e)}")
    
    def add_allocation(self):
        """Add a new project allocation"""
        try:
            dialog = ProjectAllocationDialog(self, None, datetime.now().year)
            self.wait_window(dialog)
            
            if dialog.result:
                session = get_session()
                
                # Create new allocation
                allocation = ProjectAllocation(
                    manager_code=dialog.result["manager_code"],
                    year=dialog.result["year"],
                    cost_center=dialog.result["cost_center"],
                    work_code=dialog.result["work_code"],
                    jan=dialog.result["jan"],
                    feb=dialog.result["feb"],
                    mar=dialog.result["mar"],
                    apr=dialog.result["apr"],
                    may=dialog.result["may"],
                    jun=dialog.result["jun"],
                    jul=dialog.result["jul"],
                    aug=dialog.result["aug"],
                    sep=dialog.result["sep"],
                    oct=dialog.result["oct"],
                    nov=dialog.result["nov"],
                    dec=dialog.result["dec"]
                )
                
                session.add(allocation)
                session.commit()
                session.close()
                
                # Reload allocations
                self.load_allocations()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add allocation: {str(e)}")
    
    def edit_allocation(self):
        """Edit selected project allocation"""
        try:
            selected = self.tree.selection()
            if not selected:
                messagebox.showwarning("Warning", "Please select an allocation to edit.")
                return
            
            # Get selected allocation
            values = self.tree.item(selected[0])["values"]
            
            session = get_session()
            allocation = session.query(ProjectAllocation).filter(
                ProjectAllocation.manager_code == values[0],
                ProjectAllocation.year == values[1],
                ProjectAllocation.cost_center == values[2],
                ProjectAllocation.work_code == values[3]
            ).first()
            
            if not allocation:
                session.close()
                messagebox.showerror("Error", "Selected allocation not found in database.")
                return
            
            # Open dialog with current values
            dialog = ProjectAllocationDialog(self, allocation.manager_code, allocation.year, allocation)
            self.wait_window(dialog)
            
            if dialog.result:
                # Update allocation
                allocation.manager_code = dialog.result["manager_code"]
                allocation.year = dialog.result["year"]
                allocation.cost_center = dialog.result["cost_center"]
                allocation.work_code = dialog.result["work_code"]
                allocation.jan = dialog.result["jan"]
                allocation.feb = dialog.result["feb"]
                allocation.mar = dialog.result["mar"]
                allocation.apr = dialog.result["apr"]
                allocation.may = dialog.result["may"]
                allocation.jun = dialog.result["jun"]
                allocation.jul = dialog.result["jul"]
                allocation.aug = dialog.result["aug"]
                allocation.sep = dialog.result["sep"]
                allocation.oct = dialog.result["oct"]
                allocation.nov = dialog.result["nov"]
                allocation.dec = dialog.result["dec"]
                
                session.commit()
            
            session.close()
            
            # Reload allocations
            self.load_allocations()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit allocation: {str(e)}")
    
    def delete_allocation(self):
        """Delete selected project allocation"""
        try:
            selected = self.tree.selection()
            if not selected:
                messagebox.showwarning("Warning", "Please select an allocation to delete.")
                return
            
            if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this allocation?"):
                return
            
            # Get selected allocation
            values = self.tree.item(selected[0])["values"]
            
            session = get_session()
            allocation = session.query(ProjectAllocation).filter(
                ProjectAllocation.manager_code == values[0],
                ProjectAllocation.year == values[1],
                ProjectAllocation.cost_center == values[2],
                ProjectAllocation.work_code == values[3]
            ).first()
            
            if allocation:
                session.delete(allocation)
                session.commit()
            
            session.close()
            
            # Reload allocations
            self.load_allocations()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete allocation: {str(e)}")

class ProjectAllocationDialog(tk.Toplevel):
    def __init__(self, parent, manager_code, year, allocation=None):
        super().__init__(parent)
        self.parent = parent
        self.manager_code = manager_code
        self.year = year
        self.allocation = allocation
        self.result = None
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        self.title("Project Allocation")
        self.geometry("800x600")
        self.resizable(True, True)
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Manager and Year info
        info_frame = ttk.LabelFrame(frame, text="Allocation Info", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Manager selection
        ttk.Label(info_frame, text="Manager Code:").pack(side=tk.LEFT, padx=5)
        self.manager_var = tk.StringVar(value=manager_code if manager_code else "")
        self.manager_combo = ttk.Combobox(info_frame, textvariable=self.manager_var, width=20, state="readonly")
        self.manager_combo.pack(side=tk.LEFT, padx=5)
        
        # Year selection
        ttk.Label(info_frame, text="Year:").pack(side=tk.LEFT, padx=5)
        self.year_var = tk.StringVar(value=str(year))
        years = [str(y) for y in range(datetime.now().year - 2, datetime.now().year + 3)]
        self.year_combo = ttk.Combobox(info_frame, textvariable=self.year_var, values=years, width=6, state="readonly")
        self.year_combo.pack(side=tk.LEFT, padx=5)
        
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
        
        # Create a frame for the spinboxes
        spinbox_frame = ttk.Frame(allocation_frame)
        spinbox_frame.pack(fill=tk.X, pady=5)
        
        for i, (month_name, month_code) in enumerate(months):
            row = i // 3
            col = i % 3
            
            month_frame = ttk.Frame(spinbox_frame)
            month_frame.grid(row=row, column=col, padx=5, pady=5, sticky=tk.W)
            
            ttk.Label(month_frame, text=f"{month_name}:").pack(side=tk.LEFT)
            spinbox = ttk.Spinbox(month_frame, from_=0, to=100, increment=0.5, width=10)
            spinbox.pack(side=tk.LEFT, padx=5)
            self.spinboxes[month_code] = spinbox
        
        # Add paste button
        paste_frame = ttk.Frame(allocation_frame)
        paste_frame.pack(fill=tk.X, pady=10)
        ttk.Button(paste_frame, text="Paste Monthly Values", command=self.paste_values).pack(side=tk.LEFT, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Load manager codes
        session = get_session()
        managers = session.query(Employee.manager_code).distinct().all()
        self.manager_combo['values'] = [m[0] for m in managers]
        session.close()
        
        # If editing, populate fields
        if allocation:
            self.manager_var.set(allocation.manager_code)
            self.year_var.set(str(allocation.year))
            self.cost_center_var.set(allocation.cost_center)
            self.work_code_var.set(allocation.work_code)
            
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
        
        # Center the dialog
        self.center_on_parent()
    
    def paste_values(self):
        """Paste values from clipboard into monthly spinboxes"""
        try:
            clipboard = self.clipboard_get()
            values = clipboard.strip().split()
            
            # Convert and validate values
            float_values = []
            for val in values:
                try:
                    # Handle both comma and tab delimiters
                    val = val.replace(',', '').strip()
                    float_val = float(val)
                    if float_val < 0 or float_val > 100:
                        raise ValueError(f"Value {float_val} is out of range (0-100)")
                    float_values.append(float_val)
                except ValueError as e:
                    messagebox.showerror("Error", f"Invalid value: {val}\nPlease ensure all values are numbers between 0 and 100.")
                    return
            
            # Check if we have exactly 12 values
            if len(float_values) != 12:
                messagebox.showerror("Error", f"Expected 12 values, got {len(float_values)}.\nPlease copy exactly 12 monthly values.")
                return
            
            # Set the values in the spinboxes
            months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 
                     'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
            for month, value in zip(months, float_values):
                self.spinboxes[month].set(f"{value:.1f}")
            
            messagebox.showinfo("Success", "Values pasted successfully!")
            
        except tk.TclError:
            messagebox.showerror("Error", "No data in clipboard. Please copy values first.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste values: {str(e)}")
    
    def center_on_parent(self):
        """Center the dialog on its parent window"""
        self.update_idletasks()
        
        # Get parent geometry
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        parent_x = self.parent.winfo_rootx()
        parent_y = self.parent.winfo_rooty()
        
        # Calculate position
        width = self.winfo_width()
        height = self.winfo_height()
        
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2
        
        # Set position only (preserve size)
        self.geometry(f"+{x}+{y}")
    
    def on_ok(self):
        try:
            # Validate required fields
            if not self.manager_var.get():
                raise ValueError("Manager Code is required")
            if not self.year_var.get():
                raise ValueError("Year is required")
            if not self.cost_center_var.get().strip():
                raise ValueError("Cost Center is required")
            if not self.work_code_var.get().strip():
                raise ValueError("Work Code is required")
            
            # Get monthly values
            monthly_values = {}
            for month, spinbox in self.spinboxes.items():
                try:
                    monthly_values[month] = float(spinbox.get())
                except ValueError:
                    raise ValueError(f"Invalid value for {month}")
            
            self.result = {
                "manager_code": self.manager_var.get(),
                "year": int(self.year_var.get()),
                "cost_center": self.cost_center_var.get().strip(),
                "work_code": self.work_code_var.get().strip(),
                **monthly_values
            }
            self.destroy()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
    
    def on_cancel(self):
        """Cancel dialog"""
        self.result = None
        self.destroy()
    
    def get_allocations(self):
        return self.result

class ForecastVisualization(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

class ForecastApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Forecast Tool")
        self.geometry("1200x800")
        self.configure(background=COLORS['background'])
        
        # Configure styles
        configure_styles()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.employee_tab = EmployeeTab(self.notebook)
        self.notebook.add(self.employee_tab, text="Employees")
        
        self.allocation_tab = ProjectAllocationTab(self.notebook)
        self.notebook.add(self.allocation_tab, text="Allocations")
        
        self.visualization_tab = ForecastVisualization(self.notebook)
        self.notebook.add(self.visualization_tab, text="Visualization")
        
        # Create status bar
        status_frame = ttk.Frame(self, style='Card.TFrame')
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT, padx=5)
        
        # Check database connection
        if not verify_db_connection():
            messagebox.showerror("Database Error", "Failed to connect to database.")
            self.status_var.set("Database connection failed")
        else:
            self.status_var.set("Connected to database")


if __name__ == "__main__":
    # Create database tables if they don't exist
    Base.metadata.create_all(engine)
    
    # Create the application
    app = ForecastApp()
    app.mainloop()
