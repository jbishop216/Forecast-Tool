import sys
import pandas as pd
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QTableWidget, QTableWidgetItem, QPushButton, QLabel, QLineEdit, 
    QDateEdit, QComboBox, QFileDialog, QTabWidget, QMessageBox,
    QDialog, QFormLayout, QSpinBox, QDoubleSpinBox
)
from PyQt6.QtCore import Qt, QDate

from models import Employee, GA01Week, WorkCode, Forecast, get_session
from forecast_calculator import update_forecast
from import_export import import_employees, import_ga01_weeks, export_forecast

class EmployeeDialog(QDialog):
    def __init__(self, parent=None, employee=None):
        super().__init__(parent)
        self.employee = employee
        self.setWindowTitle("Employee Details")
        self.setMinimumWidth(400)
        
        layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        layout.addRow("Name:", self.name_edit)
        
        self.team_edit = QLineEdit()
        layout.addRow("Team:", self.team_edit)
        
        self.manager_code_edit = QLineEdit()
        layout.addRow("Manager Code:", self.manager_code_edit)
        
        self.cost_center_edit = QLineEdit()
        layout.addRow("Cost Center:", self.cost_center_edit)
        
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate())
        layout.addRow("Start Date:", self.start_date_edit)
        
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate().addYears(10))
        self.end_date_edit.setEnabled(False)  # Initially disabled
        layout.addRow("End Date:", self.end_date_edit)
        
        self.has_end_date = QComboBox()
        self.has_end_date.addItems(["No", "Yes"])
        self.has_end_date.currentIndexChanged.connect(self.toggle_end_date)
        layout.addRow("Has End Date:", self.has_end_date)
        
        self.emp_type_combo = QComboBox()
        self.emp_type_combo.addItems(["FTE", "Contractor"])
        layout.addRow("Employment Type:", self.emp_type_combo)
        
        # Add OK and Cancel buttons
        button_box = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_box.addWidget(self.ok_button)
        button_box.addWidget(self.cancel_button)
        layout.addRow("", button_box)
        
        self.setLayout(layout)
        
        # Populate fields if editing existing employee
        if employee:
            self.name_edit.setText(employee.name)
            self.team_edit.setText(employee.team or "")
            self.manager_code_edit.setText(employee.manager_code)
            self.cost_center_edit.setText(employee.cost_center)
            
            if employee.start_date:
                qdate = QDate(employee.start_date.year, employee.start_date.month, employee.start_date.day)
                self.start_date_edit.setDate(qdate)
            
            if employee.end_date:
                qdate = QDate(employee.end_date.year, employee.end_date.month, employee.end_date.day)
                self.end_date_edit.setDate(qdate)
                self.has_end_date.setCurrentIndex(1)
                self.end_date_edit.setEnabled(True)
            
            index = 0 if employee.employment_type == "FTE" else 1
            self.emp_type_combo.setCurrentIndex(index)
    
    def toggle_end_date(self, index):
        self.end_date_edit.setEnabled(index == 1)
    
    def get_employee_data(self):
        """Return the employee data from the dialog"""
        name = self.name_edit.text()
        team = self.team_edit.text()
        manager_code = self.manager_code_edit.text()
        cost_center = self.cost_center_edit.text()
        
        start_date = self.start_date_edit.date().toPyDate()
        
        end_date = None
        if self.has_end_date.currentIndex() == 1:
            end_date = self.end_date_edit.date().toPyDate()
        
        employment_type = self.emp_type_combo.currentText()
        
        return {
            "name": name,
            "team": team,
            "manager_code": manager_code,
            "cost_center": cost_center,
            "start_date": start_date,
            "end_date": end_date,
            "employment_type": employment_type
        }

class GA01WeekDialog(QDialog):
    def __init__(self, parent=None, year=None):
        super().__init__(parent)
        self.setWindowTitle("GA01 Weeks")
        self.setMinimumWidth(300)
        
        if year is None:
            year = datetime.now().year
        self.year = year
        
        layout = QVBoxLayout()
        
        # Add form layout for each month
        form_layout = QFormLayout()
        
        self.spinboxes = {}
        
        for month in range(1, 13):
            month_name = datetime(2000, month, 1).strftime('%B')
            spinbox = QDoubleSpinBox()
            spinbox.setMinimum(0)
            spinbox.setMaximum(5)
            spinbox.setSingleStep(0.01)
            spinbox.setDecimals(2)
            spinbox.setValue(4.0)  # Default value
            self.spinboxes[month] = spinbox
            form_layout.addRow(f"{month_name}:", spinbox)
        
        layout.addLayout(form_layout)
        
        # Add OK and Cancel buttons
        button_box = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_box.addWidget(self.ok_button)
        button_box.addWidget(self.cancel_button)
        layout.addLayout(button_box)
        
        self.setLayout(layout)
        
        # Load GA01 weeks from database
        self.load_ga01_weeks()
    
    def load_ga01_weeks(self):
        """Load GA01 weeks from database"""
        session = get_session()
        ga01_weeks = session.query(GA01Week).filter(GA01Week.year == self.year).all()
        
        for week in ga01_weeks:
            if week.month in self.spinboxes:
                self.spinboxes[week.month].setValue(week.weeks)
        
        session.close()
    
    def get_ga01_weeks(self):
        """Return the GA01 weeks data from the dialog"""
        weeks = {}
        for month, spinbox in self.spinboxes.items():
            weeks[month] = spinbox.value()
        return weeks

class ForecastApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Employee Forecasting Tool")
        self.setMinimumSize(1000, 600)
        
        # Set up the central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create tabs
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Forecast Tab
        forecast_tab = QWidget()
        forecast_layout = QVBoxLayout(forecast_tab)
        
        # Controls for forecast view
        controls_layout = QHBoxLayout()
        
        year_label = QLabel("Year:")
        self.year_spin = QSpinBox()
        self.year_spin.setMinimum(2000)
        self.year_spin.setMaximum(2100)
        self.year_spin.setValue(datetime.now().year)
        self.year_spin.valueChanged.connect(self.load_forecast_data)
        controls_layout.addWidget(year_label)
        controls_layout.addWidget(self.year_spin)
        
        filter_label = QLabel("Filter by Manager Code:")
        self.manager_filter = QComboBox()
        self.manager_filter.addItem("All")
        self.manager_filter.currentIndexChanged.connect(self.load_forecast_data)
        controls_layout.addWidget(filter_label)
        controls_layout.addWidget(self.manager_filter)
        
        filter_cc_label = QLabel("Filter by Cost Center:")
        self.cost_center_filter = QComboBox()
        self.cost_center_filter.addItem("All")
        self.cost_center_filter.currentIndexChanged.connect(self.load_forecast_data)
        controls_layout.addWidget(filter_cc_label)
        controls_layout.addWidget(self.cost_center_filter)
        
        refresh_button = QPushButton("Refresh")
        refresh_button.clicked.connect(self.load_forecast_data)
        controls_layout.addWidget(refresh_button)
        
        forecast_layout.addLayout(controls_layout)
        
        # Forecast table
        self.forecast_table = QTableWidget()
        self.forecast_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        forecast_layout.addWidget(self.forecast_table)
        
        # Export button
        export_button = QPushButton("Export Forecast")
        export_button.clicked.connect(self.export_forecast)
        forecast_layout.addWidget(export_button)
        
        # Add forecast tab
        self.tabs.addTab(forecast_tab, "Forecast")
        
        # Employees Tab
        employees_tab = QWidget()
        employees_layout = QVBoxLayout(employees_tab)
        
        # Employees table
        self.employees_table = QTableWidget()
        self.employees_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        employees_layout.addWidget(self.employees_table)
        
        # Employee controls
        emp_controls = QHBoxLayout()
        
        add_emp_button = QPushButton("Add Employee")
        add_emp_button.clicked.connect(self.add_employee)
        emp_controls.addWidget(add_emp_button)
        
        edit_emp_button = QPushButton("Edit Employee")
        edit_emp_button.clicked.connect(self.edit_employee)
        emp_controls.addWidget(edit_emp_button)
        
        remove_emp_button = QPushButton("Remove Employee")
        remove_emp_button.clicked.connect(self.remove_employee)
        emp_controls.addWidget(remove_emp_button)
        
        import_emp_button = QPushButton("Import Employees")
        import_emp_button.clicked.connect(self.import_employees)
        emp_controls.addWidget(import_emp_button)
        
        employees_layout.addLayout(emp_controls)
        
        # Add employees tab
        self.tabs.addTab(employees_tab, "Employees")
        
        # GA01 Weeks Tab
        ga01_tab = QWidget()
        ga01_layout = QVBoxLayout(ga01_tab)
        
        # GA01 table
        self.ga01_table = QTableWidget()
        self.ga01_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        ga01_layout.addWidget(self.ga01_table)
        
        # GA01 controls
        ga01_controls = QHBoxLayout()
        
        edit_ga01_button = QPushButton("Edit GA01 Weeks")
        edit_ga01_button.clicked.connect(self.edit_ga01_weeks)
        ga01_controls.addWidget(edit_ga01_button)
        
        import_ga01_button = QPushButton("Import GA01 Weeks")
        import_ga01_button.clicked.connect(self.import_ga01_weeks)
        ga01_controls.addWidget(import_ga01_button)
        
        ga01_layout.addLayout(ga01_controls)
        
        # Add GA01 tab
        self.tabs.addTab(ga01_tab, "GA01 Weeks")
        
        # Initialize data
        self.load_employees()
        self.load_ga01_weeks()
        self.load_forecast_data()
    
    def load_employees(self):
        """Load employees data into the employees table"""
        session = get_session()
        employees = session.query(Employee).all()
        
        self.employees_table.setRowCount(len(employees))
        self.employees_table.setColumnCount(7)
        self.employees_table.setHorizontalHeaderLabels([
            "ID", "Name", "Team", "Manager Code", "Cost Center", 
            "Start Date", "End Date", "Type"
        ])
        
        # Also populate manager code and cost center filters
        manager_codes = set()
        cost_centers = set()
        
        for row, emp in enumerate(employees):
            self.employees_table.setItem(row, 0, QTableWidgetItem(str(emp.id)))
            self.employees_table.setItem(row, 1, QTableWidgetItem(emp.name))
            self.employees_table.setItem(row, 2, QTableWidgetItem(emp.team or ""))
            self.employees_table.setItem(row, 3, QTableWidgetItem(emp.manager_code))
            self.employees_table.setItem(row, 4, QTableWidgetItem(emp.cost_center))
            
            start_date = emp.start_date.strftime('%Y-%m-%d') if emp.start_date else ""
            self.employees_table.setItem(row, 5, QTableWidgetItem(start_date))
            
            end_date = emp.end_date.strftime('%Y-%m-%d') if emp.end_date else ""
            self.employees_table.setItem(row, 6, QTableWidgetItem(end_date))
            
            self.employees_table.setItem(row, 7, QTableWidgetItem(emp.employment_type))
            
            # Add to filter sets
            manager_codes.add(emp.manager_code)
            cost_centers.add(emp.cost_center)
        
        session.close()
        
        # Update filters while preserving selection if possible
        current_manager = self.manager_filter.currentText()
        self.manager_filter.clear()
        self.manager_filter.addItem("All")
        for code in sorted(manager_codes):
            self.manager_filter.addItem(code)
        
        if current_manager in manager_codes:
            index = self.manager_filter.findText(current_manager)
            if index >= 0:
                self.manager_filter.setCurrentIndex(index)
        
        current_cc = self.cost_center_filter.currentText()
        self.cost_center_filter.clear()
        self.cost_center_filter.addItem("All")
        for cc in sorted(cost_centers):
            self.cost_center_filter.addItem(cc)
        
        if current_cc in cost_centers:
            index = self.cost_center_filter.findText(current_cc)
            if index >= 0:
                self.cost_center_filter.setCurrentIndex(index)
    
    def load_ga01_weeks(self):
        """Load GA01 weeks data into the GA01 table"""
        session = get_session()
        year = datetime.now().year
        ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
        
        self.ga01_table.setRowCount(len(ga01_weeks))
        self.ga01_table.setColumnCount(3)
        self.ga01_table.setHorizontalHeaderLabels(["Year", "Month", "Weeks"])
        
        for row, week in enumerate(ga01_weeks):
            self.ga01_table.setItem(row, 0, QTableWidgetItem(str(week.year)))
            
            month_name = datetime(2000, week.month, 1).strftime('%B')
            self.ga01_table.setItem(row, 1, QTableWidgetItem(month_name))
            
            self.ga01_table.setItem(row, 2, QTableWidgetItem(str(week.weeks)))
        
        session.close()
    
    def load_forecast_data(self):
        """Load forecast data into the forecast table"""
        session = get_session()
        year = self.year_spin.value()
        
        # Apply filters
        manager_filter = self.manager_filter.currentText()
        cost_center_filter = self.cost_center_filter.currentText()
        
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
                query = query.filter(False)
        
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
        
        # Set up table
        rows = len(forecast_data)
        self.forecast_table.setRowCount(rows)
        self.forecast_table.setColumnCount(15)  # Manager, Cost Center, Work Code, 12 months, Total
        
        # Set headers
        headers = ["Manager", "Cost Center", "Work Code"]
        for month in range(1, 13):
            month_name = datetime(2000, month, 1).strftime('%b')
            headers.append(month_name)
        headers.append("Total")
        
        self.forecast_table.setHorizontalHeaderLabels(headers)
        
        # Fill table
        for row, (key, data) in enumerate(forecast_data.items()):
            self.forecast_table.setItem(row, 0, QTableWidgetItem(data["manager_code"]))
            self.forecast_table.setItem(row, 1, QTableWidgetItem(data["cost_center"]))
            self.forecast_table.setItem(row, 2, QTableWidgetItem(data["work_code"]))
            
            total = 0
            for month in range(1, 13):
                hours = data["months"].get(month, 0)
                total += hours
                self.forecast_table.setItem(row, 2 + month, QTableWidgetItem(f"{hours:.2f}"))
            
            self.forecast_table.setItem(row, 15, QTableWidgetItem(f"{total:.2f}"))
        
        session.close()
    
    def add_employee(self):
        """Add a new employee"""
        dialog = EmployeeDialog(self)
        if dialog.exec():
            data = dialog.get_employee_data()
            
            # Validate data
            if not data["name"] or not data["manager_code"] or not data["cost_center"]:
                QMessageBox.warning(self, "Missing Data", "Name, Manager Code, and Cost Center are required.")
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
        selected_row = self.employees_table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "No Selection", "Please select an employee to edit.")
            return
        
        employee_id = int(self.employees_table.item(selected_row, 0).text())
        
        session = get_session()
        employee = session.query(Employee).filter(Employee.id == employee_id).first()
        
        if not employee:
            QMessageBox.warning(self, "Error", "Employee not found.")
            session.close()
            return
        
        dialog = EmployeeDialog(self, employee)
        if dialog.exec():
            data = dialog.get_employee_data()
            
            # Validate data
            if not data["name"] or not data["manager_code"] or not data["cost_center"]:
                QMessageBox.warning(self, "Missing Data", "Name, Manager Code, and Cost Center are required.")
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
        selected_row = self.employees_table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "No Selection", "Please select an employee to remove.")
            return
        
        employee_id = int(self.employees_table.item(selected_row, 0).text())
        employee_name = self.employees_table.item(selected_row, 1).text()
        
        # Confirmation dialog
        reply = QMessageBox.question(
            self, "Confirm Deletion",
            f"Are you sure you want to delete employee '{employee_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
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
        if dialog.exec():
            weeks_data = dialog.get_ga01_weeks()
            
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
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import Employees", "", "CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        try:
            import_employees(file_path)
            QMessageBox.information(self, "Success", "Employees imported successfully.")
            self.load_employees()
            self.load_forecast_data()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import employees: {str(e)}")
    
    def import_ga01_weeks(self):
        """Import GA01 weeks from file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import GA01 Weeks", "", "CSV Files (*.csv);;Excel Files (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        try:
            year = datetime.now().year
            import_ga01_weeks(file_path, year)
            QMessageBox.information(self, "Success", "GA01 weeks imported successfully.")
            self.load_ga01_weeks()
            self.load_forecast_data()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import GA01 weeks: {str(e)}")
    
    def export_forecast(self):
        """Export forecast to file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Forecast", "", "CSV Files (*.csv);;Excel Files (*.xlsx)"
        )
        
        if not file_path:
            return
        
        try:
            year = self.year_spin.value()
            export_forecast(file_path, year)
            QMessageBox.information(self, "Success", f"Forecast exported successfully to {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export forecast: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ForecastApp()
    window.show()
    sys.exit(app.exec()) 