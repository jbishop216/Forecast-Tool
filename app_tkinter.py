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

# Define modern color scheme with better cross-platform readability
COLORS = {
    'primary': '#2c3e50',      # Dark blue-gray
    'secondary': '#2980b9',    # Darker blue for better contrast
    'accent': '#c0392b',       # Darker red for better contrast
    'success': '#27ae60',      # Darker green for better contrast
    'warning': '#f39c12',      # Darker yellow for better contrast
    'background': '#f5f5f5',   # Lighter gray for better contrast with text
    'text': '#000000',         # Black for maximum readability
    'text_light': '#2c3e50',   # Darker blue-gray for better contrast
    'white': '#ffffff',        # White
    'border': '#95a5a6',        # Darker medium gray for better visibility
    'dark': '#7f8c8d',          # Darker gray
    'table_header': '#d5d5d5',  # Light gray for table headers
    'table_row_alt': '#eaeaea'  # Alternate row color for better readability
}

# Configure ttk styles
def configure_styles():
    style = ttk.Style()
    
    # Configure main window style
    style.configure('Main.TFrame', background=COLORS['background'])
    
    # Configure notebook style with better contrast for selected tabs
    style.configure('TNotebook', background=COLORS['background'])
    style.configure('TNotebook.Tab', padding=[10, 5], background=COLORS['background'], foreground=COLORS['text'])
    style.map('TNotebook.Tab',
        background=[('selected', COLORS['secondary'])],  # Use secondary color for better visibility
        foreground=[('selected', COLORS['white'])],      # White text on selected tab
        relief=[('selected', 'sunken')]                  # Add relief effect to make selection more obvious
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
    
    # Configure treeview styles with improved contrast
    style.configure('Treeview',
        background=COLORS['white'],
        foreground=COLORS['text'],
        fieldbackground=COLORS['white'],
        rowheight=25
    )
    
    # Configure treeview headings with more visible styling
    style.configure('Treeview.Heading',
        background=COLORS['table_header'],  # Light gray background for better text contrast
        foreground=COLORS['text'],  # Black text for maximum readability
        padding=5,
        relief='raised',  # Add relief to make headings stand out
        borderwidth=1,
        font=('Helvetica', 10, 'bold')  # Make headings bold for better visibility
    )
    
    # Ensure the heading style is properly applied by explicitly setting it
    style.layout('Treeview.Heading', [
        ('Treeview.Heading.cell', {
            'sticky': 'nswe',
            'children': [
                ('Treeview.Heading.border', {
                    'sticky': 'nswe',
                    'border': '1',
                    'children': [
                        ('Treeview.Heading.padding', {
                            'sticky': 'nswe',
                            'children': [
                                ('Treeview.Heading.image', {'side': 'right', 'sticky': ''}),
                                ('Treeview.Heading.text', {'sticky': 'we'})
                            ]
                        })
                    ]
                })
            ]
        })
    ])
    
    # Add hover effect for headings
    style.map('Treeview.Heading',
        background=[('active', '#b0b0b0')],  # Slightly darker gray when active
        foreground=[('active', '#000000')]
    )
    style.map('Treeview',
        background=[('selected', COLORS['secondary'])],
        foreground=[('selected', COLORS['white'])]
    )
    
    # Add alternating row colors for better readability - handled in the treeview creation
    # as ttk.Treeview doesn't directly support alternating row colors through style mapping
    
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
        ttk.Button(button_container, text="Import Employees", command=self.import_employees).pack(side=tk.LEFT, padx=5)
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
            
            # For alternating row colors
            count = 0
            
            for emp in employees:
                item_id = self.tree.insert("", tk.END, values=(
                    emp.id,
                    emp.name,
                    emp.manager_code,
                    emp.cost_center,
                    emp.employment_type,
                    emp.start_date.strftime("%m/%d/%y") if emp.start_date else "",
                    emp.end_date.strftime("%m/%d/%y") if emp.end_date else ""
                ))
                
                # Apply alternating row colors
                if count % 2 == 1:
                    self.tree.item(item_id, tags=('evenrow',))
                else:
                    self.tree.item(item_id, tags=('oddrow',))
                count += 1
            
            # Configure row tags
            self.tree.tag_configure('oddrow', background=COLORS['white'])
            self.tree.tag_configure('evenrow', background=COLORS['table_row_alt'])
            
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
            
    def import_employees(self):
        """Import employees from a CSV file"""
        # Ask user to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return  # User cancelled
        
        try:
            # Read CSV file
            with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                
                # Check required columns
                required_columns = ['name', 'manager_code', 'cost_center', 'employment_type', 'start_date']
                headers = reader.fieldnames
                
                if not all(col in headers for col in required_columns):
                    missing = [col for col in required_columns if col not in headers]
                    messagebox.showerror("Error", f"CSV file is missing required columns: {', '.join(missing)}\n\n"
                                       f"Required columns are: {', '.join(required_columns)}")
                    return
                
                # Begin import
                session = get_session()
                imported_count = 0
                error_count = 0
                error_messages = []
                
                for row in reader:
                    try:
                        # Parse dates
                        try:
                            start_date = datetime.strptime(row['start_date'], "%m/%d/%y").date()
                        except ValueError:
                            raise ValueError(f"Invalid start date format for {row['name']}: {row['start_date']}. Use MM/DD/YY.")
                        
                        end_date = None
                        if 'end_date' in row and row['end_date']:
                            try:
                                end_date = datetime.strptime(row['end_date'], "%m/%d/%y").date()
                            except ValueError:
                                raise ValueError(f"Invalid end date format for {row['name']}: {row['end_date']}. Use MM/DD/YY.")
                        
                        # Create employee
                        employee = Employee(
                            name=row['name'],
                            manager_code=row['manager_code'],
                            cost_center=row['cost_center'],
                            employment_type=row['employment_type'],
                            start_date=start_date,
                            end_date=end_date
                        )
                        
                        session.add(employee)
                        imported_count += 1
                    except Exception as e:
                        error_count += 1
                        error_messages.append(f"Row {reader.line_num}: {str(e)}")
                        if error_count >= 5:  # Limit error messages to avoid huge dialog
                            error_messages.append("(Additional errors not shown)")
                            break
                
                # Commit changes
                if imported_count > 0:
                    session.commit()
                
                session.close()
                
                # Show results
                if error_count > 0:
                    messagebox.showwarning("Import Results", 
                                         f"Imported {imported_count} employees with {error_count} errors.\n\n"
                                         f"Errors:\n{chr(10).join(error_messages)}")
                else:
                    messagebox.showinfo("Import Success", f"Successfully imported {imported_count} employees.")
                
                # Reload data
                self.load_employees()
                
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import employees: {str(e)}")

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
            
            # For alternating row colors
            count = 0
            
            for allocation in allocations:
                item_id = self.tree.insert("", tk.END, values=(
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
                
                # Apply alternating row colors
                if count % 2 == 1:
                    self.tree.item(item_id, tags=('evenrow',))
                else:
                    self.tree.item(item_id, tags=('oddrow',))
                count += 1
            
            # Configure row tags
            self.tree.tag_configure('oddrow', background=COLORS['white'])
            self.tree.tag_configure('evenrow', background=COLORS['table_row_alt'])
            
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
        self.create_widgets()
    
    def create_widgets(self):
        # Create main container with padding
        main_frame = ttk.Frame(self, style='Card.TFrame', padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(
            header_frame,
            text="Forecast Reports & Analytics",
            font=('Helvetica', 14, 'bold'),
            foreground=COLORS['primary']
        )
        title_label.pack(side=tk.LEFT)
        
        # Add separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(0, 20))
        
        # Control frame for options
        control_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Year selection
        year_frame = ttk.Frame(control_frame)
        year_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(
            year_frame,
            text="Year:",
            foreground=COLORS['text']
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.year_combo = ttk.Combobox(
            year_frame,
            textvariable=self.year_var,
            values=[str(year) for year in range(2020, 2031)],
            width=6,
            state='readonly'
        )
        self.year_combo.pack(side=tk.LEFT)
        
        # Chart type selection
        chart_frame = ttk.Frame(control_frame)
        chart_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(
            chart_frame,
            text="Chart Type:",
            foreground=COLORS['text']
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.chart_type_var = tk.StringVar(value="Monthly Forecast")
        self.chart_type_combo = ttk.Combobox(
            chart_frame,
            textvariable=self.chart_type_var,
            values=["Monthly Forecast", "Manager Allocation", "Employee Type Distribution", "GA01 Weeks", "Planned Changes"],
            width=20,
            state='readonly'
        )
        self.chart_type_combo.pack(side=tk.LEFT)
        
        # Generate button
        ttk.Button(
            control_frame,
            text="Generate Chart",
            style='Primary.TButton',
            command=self.generate_chart
        ).pack(side=tk.LEFT)
        
        # Add separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=20)
        
        # Chart container
        self.chart_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        self.chart_frame.pack(fill=tk.BOTH, expand=True)
        
        # Bind events
        self.year_combo.bind('<<ComboboxSelected>>', lambda e: self.generate_chart())
        self.chart_type_combo.bind('<<ComboboxSelected>>', lambda e: self.generate_chart())
    
    def generate_chart(self):
        # Clear previous chart
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        
        # Create figure with white background
        fig = Figure(figsize=(10, 6), facecolor=COLORS['white'])
        ax = fig.add_subplot(111)
        
        # Set background color
        ax.set_facecolor(COLORS['white'])
        
        # Get selected year
        year = int(self.year_var.get())
        
        # Generate selected chart type
        chart_type = self.chart_type_var.get()
        if chart_type == "Monthly Forecast":
            self._generate_monthly_forecast_chart(ax, year)
        elif chart_type == "Manager Allocation":
            self._generate_manager_allocation_chart(ax, year)
        elif chart_type == "Employee Type Distribution":
            self._generate_employee_type_distribution(ax, year)
        elif chart_type == "GA01 Weeks":
            self._generate_ga01_weeks_chart(ax, year)
        elif chart_type == "Planned Changes":
            self._generate_planned_changes_chart(ax, year)
        
        # Create canvas
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def _generate_monthly_forecast_chart(self, ax, year):
        session = get_session()
        forecasts = session.query(Forecast).filter(Forecast.year == year).all()
        session.close()
        
        if not forecasts:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        total_hours = [0] * 12
        
        for forecast in forecasts:
            total_hours[0] += forecast.jan
            total_hours[1] += forecast.feb
            total_hours[2] += forecast.mar
            total_hours[3] += forecast.apr
            total_hours[4] += forecast.may
            total_hours[5] += forecast.jun
            total_hours[6] += forecast.jul
            total_hours[7] += forecast.aug
            total_hours[8] += forecast.sep
            total_hours[9] += forecast.oct
            total_hours[10] += forecast.nov
            total_hours[11] += forecast.dec
        
        # Create bar chart
        bars = ax.bar(months, total_hours, color=COLORS['secondary'])
        
        # Customize chart
        ax.set_title(f'Monthly Forecast Hours - {year}', 
                    color=COLORS['primary'], 
                    pad=20)
        ax.set_xlabel('Month', color=COLORS['text'])
        ax.set_ylabel('Total Hours', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom')
        
        # Rotate x-axis labels for better readability
        plt.xticks(rotation=45)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_manager_allocation_chart(self, ax, year):
        session = get_session()
        forecasts = session.query(Forecast).filter(Forecast.year == year).all()
        session.close()
        
        if not forecasts:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        manager_data = {}
        for forecast in forecasts:
            if forecast.manager_code not in manager_data:
                manager_data[forecast.manager_code] = 0
            manager_data[forecast.manager_code] += forecast.total_hours
        
        # Sort managers by total hours
        sorted_managers = sorted(manager_data.items(), 
                               key=lambda x: x[1], 
                               reverse=True)
        
        # Create horizontal bar chart
        managers = [m[0] for m in sorted_managers]
        hours = [m[1] for m in sorted_managers]
        
        bars = ax.barh(managers, hours, color=COLORS['secondary'])
        
        # Customize chart
        ax.set_title(f'Manager Allocation - {year}', 
                    color=COLORS['primary'], 
                    pad=20)
        ax.set_xlabel('Total Hours', color=COLORS['text'])
        ax.set_ylabel('Manager Code', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2.,
                   f'{int(width)}',
                   ha='left', va='center', pad=5)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_employee_type_distribution(self, ax, year):
        session = get_session()
        employees = session.query(Employee).all()
        session.close()
        
        if not employees:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Count employee types
        type_counts = {}
        for employee in employees:
            if employee.employment_type not in type_counts:
                type_counts[employee.employment_type] = 0
            type_counts[employee.employment_type] += 1
        
        # Create pie chart
        types = list(type_counts.keys())
        counts = list(type_counts.values())
        
        colors = [COLORS['secondary'], COLORS['accent']]
        patches, texts, autotexts = ax.pie(counts, 
                                         labels=types,
                                         colors=colors,
                                         autopct='%1.1f%%',
                                         startangle=90)
        
        # Customize chart
        ax.set_title(f'Employee Type Distribution - {year}', 
                    color=COLORS['primary'], 
                    pad=20)
        
        # Customize text properties
        plt.setp(autotexts, size=9, weight="bold")
        plt.setp(texts, size=10)
    
    def _generate_ga01_weeks_chart(self, ax, year):
        session = get_session()
        ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
        session.close()
        
        if not ga01_weeks:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        weeks = [week.weeks for week in ga01_weeks]
        
        # Create bar chart
        bars = ax.bar(months, weeks, color=COLORS['accent'])
        
        # Customize chart
        ax.set_title(f'GA01 Weeks - {year}', 
                    color=COLORS['accent'], 
                    pad=20)
        ax.set_xlabel('Month', color=COLORS['text'])
        ax.set_ylabel('Weeks', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom')
        
        # Rotate x-axis labels for better readability
        plt.xticks(rotation=45)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_planned_changes_chart(self, ax, year):
        session = get_session()
        changes = session.query(PlannedChange).filter(PlannedChange.effective_date.between(
            datetime(year, 1, 1).date(), datetime(year, 12, 31).date())).all()
        session.close()
        
        if not changes:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        change_types = ['New Hire', 'Conversion', 'Termination']
        change_counts = [sum(1 for change in changes if change.change_type == change_type) for change_type in change_types]
        
        # Create bar chart
        bars = ax.bar(change_types, change_counts, color=COLORS['accent'])
        
        # Customize chart
        ax.set_title(f'Planned Changes - {year}', 
                    color=COLORS['accent'], 
                    pad=20)
        ax.set_xlabel('Change Type', color=COLORS['text'])
        ax.set_ylabel('Count', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom')
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_allocation_chart(self, ax, year):
        session = get_session()
        allocations = session.query(ProjectAllocation).filter(ProjectAllocation.year == year).all()
        session.close()
        
        if not allocations:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        manager_data = {}
        for allocation in allocations:
            if allocation.manager_code not in manager_data:
                manager_data[allocation.manager_code] = 0
            manager_data[allocation.manager_code] += allocation.total_hours
        
        # Sort managers by total hours
        sorted_managers = sorted(manager_data.items(), 
                               key=lambda x: x[1], 
                               reverse=True)
        
        # Create horizontal bar chart
        managers = [m[0] for m in sorted_managers]
        hours = [m[1] for m in sorted_managers]
        
        bars = ax.barh(managers, hours, color=COLORS['accent'])
        
        # Customize chart
        ax.set_title(f'Project Allocations - {year}', 
                    color=COLORS['accent'], 
                    pad=20)
        ax.set_xlabel('Total Hours', color=COLORS['text'])
        ax.set_ylabel('Manager Code', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2.,
                   f'{int(width)}',
                   ha='left', va='center', pad=5)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)

class ForecastTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Create toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Add Forecast", command=self.add_forecast).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Forecast", command=self.edit_forecast).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Forecast", command=self.delete_forecast).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_forecasts).pack(side=tk.LEFT, padx=2)
        
        # Year filter
        year_frame = ttk.Frame(toolbar)
        year_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(year_frame, text="Year:").pack(side=tk.LEFT, padx=2)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        year_combo = ttk.Combobox(
            year_frame, 
            textvariable=self.year_var,
            values=[str(y) for y in range(datetime.now().year - 2, datetime.now().year + 5)],
            width=6,
            state="readonly"
        )
        year_combo.pack(side=tk.LEFT, padx=2)
        year_combo.bind("<<ComboboxSelected>>", lambda e: self.load_forecasts())
        
        # Create treeview with scrollbar
        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=(
            "id", "manager_code", "cost_center", "work_code",
            "jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "oct", "nov", "dec", "total"
        ), show="headings")
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        self.tree.heading("id", text="ID")
        self.tree.heading("manager_code", text="Manager")
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
        self.tree.heading("total", text="Total")
        
        # Set column widths
        self.tree.column("id", width=50)
        self.tree.column("manager_code", width=80)
        self.tree.column("cost_center", width=80)
        self.tree.column("work_code", width=80)
        for month in ["jan", "feb", "mar", "apr", "may", "jun",
                     "jul", "aug", "sep", "oct", "nov", "dec"]:
            self.tree.column(month, width=50)
        self.tree.column("total", width=70)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Load forecasts
        self.load_forecasts()
    
    def load_forecasts(self):
        """Load forecasts from database"""
        try:
            # Clear existing items
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            session = get_session()
            year = int(self.year_var.get())
            forecasts = session.query(Forecast).filter(Forecast.year == year).all()
            
            for forecast in forecasts:
                total = sum([
                    forecast.jan, forecast.feb, forecast.mar, 
                    forecast.apr, forecast.may, forecast.jun,
                    forecast.jul, forecast.aug, forecast.sep,
                    forecast.oct, forecast.nov, forecast.dec
                ])
                
                self.tree.insert("", tk.END, values=(
                    forecast.id,
                    forecast.manager_code,
                    forecast.cost_center,
                    forecast.work_code,
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
                    total
                ))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load forecasts: {str(e)}")
    
    def add_forecast(self):
        """Add a new forecast"""
        dialog = ForecastDialog(self, None)
        self.wait_window(dialog)
        
        if dialog.result:
            try:
                session = get_session()
                
                # Calculate total
                total_hours = sum([
                    dialog.result["jan"], dialog.result["feb"], dialog.result["mar"],
                    dialog.result["apr"], dialog.result["may"], dialog.result["jun"],
                    dialog.result["jul"], dialog.result["aug"], dialog.result["sep"],
                    dialog.result["oct"], dialog.result["nov"], dialog.result["dec"]
                ])
                
                # Create new forecast
                forecast = Forecast(
                    year=dialog.result["year"],
                    manager_code=dialog.result["manager_code"],
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
                    dec=dialog.result["dec"],
                    total_hours=total_hours
                )
                
                session.add(forecast)
                session.commit()
                session.close()
                
                # Reload forecasts
                self.load_forecasts()
                
                messagebox.showinfo("Success", "Forecast added successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add forecast: {str(e)}")
    
    def edit_forecast(self):
        """Edit selected forecast"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a forecast to edit.")
            return
        
        try:
            # Get ID of selected forecast
            forecast_id = self.tree.item(selected[0])["values"][0]
            
            session = get_session()
            forecast = session.query(Forecast).filter(Forecast.id == forecast_id).first()
            
            if not forecast:
                session.close()
                messagebox.showerror("Error", "Selected forecast not found.")
                return
            
            # Create dialog
            dialog = ForecastDialog(self, forecast)
            self.wait_window(dialog)
            
            if dialog.result:
                # Calculate total
                total_hours = sum([
                    dialog.result["jan"], dialog.result["feb"], dialog.result["mar"],
                    dialog.result["apr"], dialog.result["may"], dialog.result["jun"],
                    dialog.result["jul"], dialog.result["aug"], dialog.result["sep"],
                    dialog.result["oct"], dialog.result["nov"], dialog.result["dec"]
                ])
                
                # Update forecast
                forecast.year = dialog.result["year"]
                forecast.manager_code = dialog.result["manager_code"]
                forecast.cost_center = dialog.result["cost_center"]
                forecast.work_code = dialog.result["work_code"]
                forecast.jan = dialog.result["jan"]
                forecast.feb = dialog.result["feb"]
                forecast.mar = dialog.result["mar"]
                forecast.apr = dialog.result["apr"]
                forecast.may = dialog.result["may"]
                forecast.jun = dialog.result["jun"]
                forecast.jul = dialog.result["jul"]
                forecast.aug = dialog.result["aug"]
                forecast.sep = dialog.result["sep"]
                forecast.oct = dialog.result["oct"]
                forecast.nov = dialog.result["nov"]
                forecast.dec = dialog.result["dec"]
                forecast.total_hours = total_hours
                
                session.commit()
                session.close()
                
                # Reload forecasts
                self.load_forecasts()
                
                messagebox.showinfo("Success", "Forecast updated successfully.")
            else:
                session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit forecast: {str(e)}")
    
    def delete_forecast(self):
        """Delete selected forecast"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a forecast to delete.")
            return
        
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this forecast?"):
            return
        
        try:
            # Get ID of selected forecast
            forecast_id = self.tree.item(selected[0])["values"][0]
            
            session = get_session()
            forecast = session.query(Forecast).filter(Forecast.id == forecast_id).first()
            
            if forecast:
                session.delete(forecast)
                session.commit()
                
                # Reload forecasts
                self.load_forecasts()
                
                messagebox.showinfo("Success", "Forecast deleted successfully.")
            else:
                messagebox.showwarning("Warning", "Selected forecast not found.")
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete forecast: {str(e)}")

class ForecastDialog(tk.Toplevel):
    def __init__(self, parent, forecast=None):
        super().__init__(parent)
        self.parent = parent
        self.forecast = forecast
        self.result = None
        
        self.title("Forecast")
        self.geometry("800x600")
        self.resizable(True, True)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Forecast info
        info_frame = ttk.LabelFrame(frame, text="Forecast Info", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Manager selection
        ttk.Label(info_frame, text="Manager Code:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.manager_var = tk.StringVar(value=forecast.manager_code if forecast else "")
        managers_combo = ttk.Combobox(info_frame, textvariable=self.manager_var, width=20)
        managers_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Year selection
        ttk.Label(info_frame, text="Year:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.year_var = tk.StringVar(value=str(forecast.year if forecast else datetime.now().year))
        years = [str(y) for y in range(datetime.now().year - 2, datetime.now().year + 5)]
        ttk.Combobox(info_frame, textvariable=self.year_var, values=years, width=6).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Cost center
        ttk.Label(info_frame, text="Cost Center:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.cost_center_var = tk.StringVar(value=forecast.cost_center if forecast else "")
        ttk.Entry(info_frame, textvariable=self.cost_center_var, width=20).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Work code
        ttk.Label(info_frame, text="Work Code:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.work_code_var = tk.StringVar(value=forecast.work_code if forecast else "")
        ttk.Entry(info_frame, textvariable=self.work_code_var, width=20).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Monthly allocations
        allocation_frame = ttk.LabelFrame(frame, text="Monthly Hours", padding="10")
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
            spinbox = ttk.Spinbox(month_frame, from_=0, to=1000, increment=0.5, width=10)
            spinbox.pack(side=tk.LEFT, padx=5)
            self.spinboxes[month_code] = spinbox
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Load manager codes
        session = get_session()
        managers = session.query(Employee.manager_code).distinct().all()
        managers_combo['values'] = [m[0] for m in managers]
        session.close()
        
        # If editing, populate fields
        if forecast:
            self.manager_var.set(forecast.manager_code)
            self.year_var.set(str(forecast.year))
            self.cost_center_var.set(forecast.cost_center)
            self.work_code_var.set(forecast.work_code)
            
            # Set monthly values
            self.spinboxes["jan"].set(forecast.jan or 0)
            self.spinboxes["feb"].set(forecast.feb or 0)
            self.spinboxes["mar"].set(forecast.mar or 0)
            self.spinboxes["apr"].set(forecast.apr or 0)
            self.spinboxes["may"].set(forecast.may or 0)
            self.spinboxes["jun"].set(forecast.jun or 0)
            self.spinboxes["jul"].set(forecast.jul or 0)
            self.spinboxes["aug"].set(forecast.aug or 0)
            self.spinboxes["sep"].set(forecast.sep or 0)
            self.spinboxes["oct"].set(forecast.oct or 0)
            self.spinboxes["nov"].set(forecast.nov or 0)
            self.spinboxes["dec"].set(forecast.dec or 0)
        
        # Center the dialog
        self.center_on_parent()
    
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

class PlannedChangesTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        # Create toolbar
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Add Change", command=self.add_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Change", command=self.edit_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Change", command=self.delete_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_changes).pack(side=tk.LEFT, padx=2)
        
        # Year filter
        year_frame = ttk.Frame(toolbar)
        year_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(year_frame, text="Year:").pack(side=tk.LEFT, padx=2)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        year_combo = ttk.Combobox(
            year_frame, 
            textvariable=self.year_var,
            values=[str(y) for y in range(datetime.now().year - 2, datetime.now().year + 5)],
            width=6,
            state="readonly"
        )
        year_combo.pack(side=tk.LEFT, padx=2)
        year_combo.bind("<<ComboboxSelected>>", lambda e: self.load_changes())
        
        # Create treeview with scrollbar
        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(self.tree_frame, columns=(
            "id", "description", "change_type", "effective_date", "name", "team", "manager_code", "status"
        ), show="headings")
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        self.tree.heading("id", text="ID")
        self.tree.heading("description", text="Description")
        self.tree.heading("change_type", text="Change Type")
        self.tree.heading("effective_date", text="Effective Date")
        self.tree.heading("name", text="Name")
        self.tree.heading("team", text="Team")
        self.tree.heading("manager_code", text="Manager")
        self.tree.heading("status", text="Status")
        
        # Set column widths
        self.tree.column("id", width=50)
        self.tree.column("description", width=200)
        self.tree.column("change_type", width=100)
        self.tree.column("effective_date", width=100)
        self.tree.column("name", width=150)
        self.tree.column("team", width=100)
        self.tree.column("manager_code", width=100)
        self.tree.column("status", width=80)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Load changes
        self.load_changes()
    
    def load_changes(self):
        """Load planned changes from database"""
        try:
            # Clear existing items
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            session = get_session()
            year = int(self.year_var.get())
            changes = session.query(PlannedChange).filter(
                PlannedChange.effective_date.between(
                    datetime(year, 1, 1).date(), 
                    datetime(year, 12, 31).date()
                )
            ).all()
            
            # For alternating row colors
            count = 0
            
            for change in changes:
                item_id = self.tree.insert("", tk.END, values=(
                    change.id,
                    change.description,
                    change.change_type,
                    change.effective_date.strftime("%m/%d/%y") if change.effective_date else "",
                    change.name or "",
                    change.team or "",
                    change.manager_code or "",
                    change.status
                ))
                
                # Apply alternating row colors
                if count % 2 == 1:
                    self.tree.item(item_id, tags=('evenrow',))
                else:
                    self.tree.item(item_id, tags=('oddrow',))
                count += 1
            
            # Configure row tags
            self.tree.tag_configure('oddrow', background=COLORS['white'])
            self.tree.tag_configure('evenrow', background=COLORS['table_row_alt'])
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load planned changes: {str(e)}")
    
    def add_change(self):
        """Add a new planned change"""
        dialog = PlannedChangeDialog(self, None)
        self.wait_window(dialog)
        
        if dialog.result:
            try:
                session = get_session()
                
                # Create new planned change
                change = PlannedChange(
                    description=dialog.result["description"],
                    change_type=dialog.result["change_type"],
                    effective_date=dialog.result["effective_date"],
                    name=dialog.result.get("name", ""),
                    team=dialog.result.get("team", ""),
                    manager_code=dialog.result.get("manager_code", ""),
                    cost_center=dialog.result.get("cost_center", ""),
                    employment_type=dialog.result.get("employment_type", ""),
                    status=dialog.result["status"]
                )
                
                if dialog.result.get("employee_id"):
                    change.employee_id = dialog.result["employee_id"]
                
                session.add(change)
                session.commit()
                session.close()
                
                # Reload changes
                self.load_changes()
                
                messagebox.showinfo("Success", "Planned change added successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add planned change: {str(e)}")
    
    def edit_change(self):
        """Edit selected planned change"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a planned change to edit.")
            return
        
        try:
            # Get ID of selected change
            change_id = self.tree.item(selected[0])["values"][0]
            
            session = get_session()
            change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
            
            if not change:
                session.close()
                messagebox.showerror("Error", "Selected planned change not found.")
                return
            
            # Create dialog
            dialog = PlannedChangeDialog(self, change)
            self.wait_window(dialog)
            
            if dialog.result:
                # Update change
                change.description = dialog.result["description"]
                change.change_type = dialog.result["change_type"]
                change.effective_date = dialog.result["effective_date"]
                change.name = dialog.result.get("name", "")
                change.team = dialog.result.get("team", "")
                change.manager_code = dialog.result.get("manager_code", "")
                change.cost_center = dialog.result.get("cost_center", "")
                change.employment_type = dialog.result.get("employment_type", "")
                change.status = dialog.result["status"]
                
                if dialog.result.get("employee_id"):
                    change.employee_id = dialog.result["employee_id"]
                else:
                    change.employee_id = None
                
                session.commit()
            
            session.close()
            
            # Reload changes
            self.load_changes()
            
            messagebox.showinfo("Success", "Planned change updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit planned change: {str(e)}")
    
    def delete_change(self):
        """Delete selected planned change"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a planned change to delete.")
            return
        
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this planned change?"):
            return
        
        try:
            # Get ID of selected change
            change_id = self.tree.item(selected[0])["values"][0]
            
            session = get_session()
            change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
            
            if change:
                session.delete(change)
                session.commit()
            
            session.close()
            
            # Reload changes
            self.load_changes()
            
            messagebox.showinfo("Success", "Planned change deleted successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete planned change: {str(e)}")

class PlannedChangeDialog(tk.Toplevel):
    def __init__(self, parent, change=None):
        super().__init__(parent)
        self.parent = parent
        self.change = change
        self.result = None
        self.employee_id = None
        if change and change.employee_id:
            self.employee_id = change.employee_id
        
        self.title("Planned Change")
        self.geometry("550x550")
        self.resizable(True, True)
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create form
        frame = ttk.Frame(self, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Change info
        ttk.Label(frame, text="Description:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.description_var = tk.StringVar(value=change.description if change else "")
        ttk.Entry(frame, textvariable=self.description_var, width=40).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Change type
        ttk.Label(frame, text="Change Type:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.change_type_var = tk.StringVar(value=change.change_type if change else "")
        change_types = [ct.value for ct in ChangeType]
        self.change_type_combo = ttk.Combobox(frame, textvariable=self.change_type_var, values=change_types, width=20, state="readonly")
        self.change_type_combo.grid(row=1, column=1, sticky=tk.W, pady=5)
        self.change_type_combo.bind("<<ComboboxSelected>>", self.on_change_type_selected)
        
        # Effective date
        ttk.Label(frame, text="Effective Date (mm/dd/yy):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.effective_date_var = tk.StringVar(value=change.effective_date.strftime("%m/%d/%y") if change and change.effective_date else datetime.now().strftime("%m/%d/%y"))
        ttk.Entry(frame, textvariable=self.effective_date_var, width=20).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Employee selection frame for termination/conversion
        self.employee_selection_frame = ttk.LabelFrame(frame, text="Select Employee", padding=10)
        self.employee_selection_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky=tk.EW)
        self.employee_selection_frame.grid_remove()  # Hide initially
        
        # Employee listbox with scrollbar
        self.employee_listbox_frame = ttk.Frame(self.employee_selection_frame)
        self.employee_listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.employee_listbox = tk.Listbox(self.employee_listbox_frame, height=6, width=40)
        self.employee_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(self.employee_listbox_frame, orient=tk.VERTICAL, command=self.employee_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.employee_listbox.config(yscrollcommand=scrollbar.set)
        
        # Bind selection event
        self.employee_listbox.bind('<<ListboxSelect>>', self.on_employee_selected)
        
        # Employee details frame
        self.details_frame = ttk.LabelFrame(frame, text="Employee Details", padding=10)
        self.details_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky=tk.EW)
        
        # Name
        ttk.Label(self.details_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar(value=change.name if change else "")
        self.name_entry = ttk.Entry(self.details_frame, textvariable=self.name_var, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Team
        ttk.Label(self.details_frame, text="Team:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.team_var = tk.StringVar(value=change.team if change else "")
        self.team_entry = ttk.Entry(self.details_frame, textvariable=self.team_var, width=30)
        self.team_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Manager code
        ttk.Label(self.details_frame, text="Manager Code:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.manager_code_var = tk.StringVar(value=change.manager_code if change else "")
        self.manager_code_entry = ttk.Entry(self.details_frame, textvariable=self.manager_code_var, width=20)
        self.manager_code_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Cost center
        ttk.Label(self.details_frame, text="Cost Center:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.cost_center_var = tk.StringVar(value=change.cost_center if change else "")
        self.cost_center_entry = ttk.Entry(self.details_frame, textvariable=self.cost_center_var, width=20)
        self.cost_center_entry.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Employment type
        ttk.Label(self.details_frame, text="Employment Type:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.employment_type_var = tk.StringVar(value=change.employment_type if change else "")
        employment_types = [et.value for et in EmploymentType]
        self.employment_type_combo = ttk.Combobox(self.details_frame, textvariable=self.employment_type_var, values=employment_types, width=20)
        self.employment_type_combo.grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # Status
        ttk.Label(frame, text="Status:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.status_var = tk.StringVar(value=change.status if change else "Planned")
        statuses = ["Planned", "In Progress", "Completed", "Cancelled"]
        ttk.Combobox(frame, textvariable=self.status_var, values=statuses, width=20, state="readonly").grid(row=5, column=1, sticky=tk.W, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="OK", command=self.on_ok, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)
        
        # Center the dialog
        self.center_on_parent()
        
        # Load employees for listbox if change type is already set
        if change and change.change_type in [ChangeType.CONVERSION.value, ChangeType.TERMINATION.value]:
            self.on_change_type_selected(None)
            
            # Set the selected employee if there is one
            if change.employee_id:
                self.employee_id = change.employee_id
                self.load_employee_details(change.employee_id)
    
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
    
    def on_change_type_selected(self, event):
        """Handle change type selection"""
        change_type = self.change_type_var.get()
        
        if change_type in [ChangeType.CONVERSION.value, ChangeType.TERMINATION.value]:
            # Show employee selection for conversion/termination
            self.employee_selection_frame.grid()
            self.load_employees()
        else:
            # Hide employee selection for other change types
            self.employee_selection_frame.grid_remove()
            self.employee_id = None
            
            # Clear and enable employee detail fields
            self.name_var.set("")
            self.team_var.set("")
            self.manager_code_var.set("")
            self.cost_center_var.set("")
            self.employment_type_var.set("")
            
            self.name_entry.config(state="normal")
            self.team_entry.config(state="normal")
            self.manager_code_entry.config(state="normal")
            self.cost_center_entry.config(state="normal")
            self.employment_type_combo.config(state="normal")
    
    def load_employees(self):
        """Load employees into the listbox"""
        try:
            # Clear the listbox
            self.employee_listbox.delete(0, tk.END)
            
            # Get employees from database
            session = get_session()
            employees = session.query(Employee).all()
            
            # Store employee data for later use
            self.employees = {}
            
            # Add employees to listbox
            for emp in employees:
                display_text = f"{emp.name} - {emp.manager_code} ({emp.employment_type})"
                self.employee_listbox.insert(tk.END, display_text)
                # Store employee data with the index
                index = self.employee_listbox.size() - 1
                self.employees[index] = emp
                
                # Select the employee if it matches the current change
                if self.employee_id and emp.id == self.employee_id:
                    self.employee_listbox.selection_set(index)
                    self.employee_listbox.see(index)
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employees: {str(e)}")
    
    def on_employee_selected(self, event):
        """Handle employee selection from listbox"""
        selection = self.employee_listbox.curselection()
        if selection:
            index = selection[0]
            employee = self.employees[index]
            self.employee_id = employee.id
            self.load_employee_details(employee.id)
    
    def load_employee_details(self, employee_id):
        """Load details of the selected employee"""
        try:
            session = get_session()
            employee = session.query(Employee).filter(Employee.id == employee_id).first()
            
            if employee:
                # Set employee details in form
                self.name_var.set(employee.name)
                self.team_var.set("")  # Assuming team is not in Employee model
                self.manager_code_var.set(employee.manager_code)
                self.cost_center_var.set(employee.cost_center)
                self.employment_type_var.set(employee.employment_type)
                
                # Disable editing of employee details for termination/conversion
                self.name_entry.config(state="readonly")
                self.team_entry.config(state="readonly")
                self.manager_code_entry.config(state="readonly")
                self.cost_center_entry.config(state="readonly")
                self.employment_type_combo.config(state="readonly")
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employee details: {str(e)}")
    
    def on_ok(self):
        try:
            # Validate required fields
            if not self.description_var.get().strip():
                raise ValueError("Description is required")
            if not self.change_type_var.get():
                raise ValueError("Change Type is required")
            if not self.effective_date_var.get().strip():
                raise ValueError("Effective Date is required")
            if not self.status_var.get():
                raise ValueError("Status is required")
                
            # Validate employee selection for termination/conversion
            if self.change_type_var.get() in [ChangeType.CONVERSION.value, ChangeType.TERMINATION.value] and not self.employee_id:
                raise ValueError("Please select an employee for termination or conversion")
            
            # Parse effective date
            try:
                effective_date = datetime.strptime(self.effective_date_var.get(), "%m/%d/%y").date()
            except ValueError:
                raise ValueError("Invalid Effective Date format. Use MM/DD/YY.")
            
            # Collect result
            self.result = {
                "description": self.description_var.get().strip(),
                "change_type": self.change_type_var.get(),
                "effective_date": effective_date,
                "name": self.name_var.get().strip(),
                "team": self.team_var.get().strip(),
                "manager_code": self.manager_code_var.get().strip(),
                "cost_center": self.cost_center_var.get().strip(),
                "employment_type": self.employment_type_var.get(),
                "status": self.status_var.get()
            }
            
            # Include employee ID if selected
            if self.employee_id:
                self.result["employee_id"] = self.employee_id
            
            # If this is an existing change, include the ID
            if self.change and self.change.id:
                self.result["id"] = self.change.id
            
            self.destroy()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
    
    def on_cancel(self):
        """Cancel dialog"""
        self.result = None
        self.destroy()

class SettingsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="20")
        self.parent = parent
        
        # Create main frame with some padding
        main_frame = ttk.Frame(self, style='Card.TFrame', padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_frame, text="Application Settings", 
                               font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Settings form
        settings_frame = ttk.Frame(main_frame)
        settings_frame.pack(fill=tk.X, pady=10)
        
        # FTE Hours
        fte_frame = ttk.Frame(settings_frame)
        fte_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(fte_frame, text="FTE Hours per Week:", width=20).pack(side=tk.LEFT, padx=5)
        self.fte_hours_var = tk.StringVar()
        fte_entry = ttk.Entry(fte_frame, textvariable=self.fte_hours_var, width=10)
        fte_entry.pack(side=tk.LEFT, padx=5)
        
        # Contractor Hours
        contractor_frame = ttk.Frame(settings_frame)
        contractor_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(contractor_frame, text="Contractor Hours per Week:", width=20).pack(side=tk.LEFT, padx=5)
        self.contractor_hours_var = tk.StringVar()
        contractor_entry = ttk.Entry(contractor_frame, textvariable=self.contractor_hours_var, width=10)
        contractor_entry.pack(side=tk.LEFT, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Save Settings", command=self.save_settings, 
                  style="Primary.TButton", width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset to Defaults", command=self.reset_defaults, 
                  width=15).pack(side=tk.LEFT, padx=5)
        
        # Load current settings
        self.load_settings()
    
    def load_settings(self):
        """Load settings from database"""
        try:
            session = get_session()
            settings = session.query(Settings).first()
            
            if not settings:
                # Create default settings if none exist
                settings = Settings(fte_hours=34.5, contractor_hours=39.0)
                session.add(settings)
                session.commit()
            
            # Set values in form
            self.fte_hours_var.set(str(settings.fte_hours))
            self.contractor_hours_var.set(str(settings.contractor_hours))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load settings: {str(e)}")
    
    def save_settings(self):
        """Save settings to database"""
        try:
            # Validate inputs
            try:
                fte_hours = float(self.fte_hours_var.get())
                contractor_hours = float(self.contractor_hours_var.get())
                
                if fte_hours <= 0 or contractor_hours <= 0:
                    raise ValueError("Hours must be positive numbers")
                    
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers for hours")
                return
            
            # Save to database
            session = get_session()
            settings = session.query(Settings).first()
            
            if not settings:
                settings = Settings()
                session.add(settings)
            
            settings.fte_hours = fte_hours
            settings.contractor_hours = contractor_hours
            
            session.commit()
            session.close()
            
            messagebox.showinfo("Success", "Settings saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def reset_defaults(self):
        """Reset settings to default values"""
        self.fte_hours_var.set("34.5")
        self.contractor_hours_var.set("39.0")

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
        
        self.forecast_tab = ForecastTab(self.notebook)
        self.notebook.add(self.forecast_tab, text="Forecast")
        
        self.planned_changes_tab = PlannedChangesTab(self.notebook) 
        self.notebook.add(self.planned_changes_tab, text="Planned Changes")
        
        self.visualization_tab = ForecastVisualization(self.notebook)
        self.notebook.add(self.visualization_tab, text="Visualization")
        
        self.settings_tab = SettingsTab(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        
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
