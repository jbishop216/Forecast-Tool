import pandas as pd
from models import Employee, GA01Week, WorkCode, Forecast, get_session
from datetime import datetime
from forecast_calculator import update_forecast

def import_employees(file_path):
    """Import employees from CSV or Excel file"""
    session = get_session()
    
    # Detect file type based on extension
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please use CSV or Excel.")

    # Expected columns: name, team, manager_code, cost_center, start_date, end_date, employment_type
    required_columns = ['name', 'manager_code', 'cost_center', 'start_date', 'employment_type']
    
    # Check for required columns
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
    
    # Process each row
    for _, row in df.iterrows():
        # Parse dates
        start_date = pd.to_datetime(row['start_date']).date() if pd.notna(row['start_date']) else None
        end_date = pd.to_datetime(row['end_date']).date() if 'end_date' in row and pd.notna(row['end_date']) else None
        
        # Create employee
        employee = Employee(
            name=row['name'],
            team=row['team'] if 'team' in row and pd.notna(row['team']) else None,
            manager_code=row['manager_code'],
            cost_center=row['cost_center'],
            start_date=start_date,
            end_date=end_date,
            employment_type=row['employment_type']
        )
        
        session.add(employee)
    
    session.commit()
    session.close()
    
    # Update forecasts for all employees
    update_forecast()

def import_ga01_weeks(file_path, year=None):
    """Import GA01 weeks from CSV or Excel file"""
    session = get_session()
    
    # Detect file type based on extension
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    elif file_path.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please use CSV or Excel.")
    
    # If year not provided, use current year
    if year is None:
        year = datetime.now().year
    
    # Expected columns: month, weeks
    required_columns = ['month', 'weeks']
    
    # Check for required columns
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
    
    # Delete existing GA01 weeks for this year
    session.query(GA01Week).filter(GA01Week.year == year).delete()
    
    # Process each row
    for _, row in df.iterrows():
        month = int(row['month'])
        weeks = float(row['weeks'])
        
        if month < 1 or month > 12:
            raise ValueError(f"Invalid month: {month}. Month must be between 1 and 12.")
        
        ga01 = GA01Week(
            year=year,
            month=month,
            weeks=weeks
        )
        
        session.add(ga01)
    
    session.commit()
    session.close()
    
    # Update forecasts for all employees
    update_forecast()

def export_forecast(file_path, year=None):
    """Export forecast to CSV or Excel file"""
    session = get_session()
    
    # If year not provided, use current year
    if year is None:
        year = datetime.now().year
    
    # Get forecasts for this year
    forecasts = session.query(Forecast).filter(Forecast.year == year).all()
    
    # Create a list to store the data
    data = []
    
    for forecast in forecasts:
        employee = session.query(Employee).filter(Employee.id == forecast.employee_id).first()
        work_code = session.query(WorkCode).filter(WorkCode.id == forecast.work_code_id).first()
        
        data.append({
            'Employee': employee.name,
            'Team': employee.team,
            'Manager Code': employee.manager_code,
            'Cost Center': employee.cost_center,
            'Work Code': work_code.code,
            'Month': forecast.month,
            'Hours': forecast.hours,
            'Employment Type': employee.employment_type
        })
    
    session.close()
    
    # Convert to DataFrame
    df = pd.DataFrame(data)
    
    # Save to file
    if file_path.endswith('.csv'):
        df.to_csv(file_path, index=False)
    elif file_path.endswith('.xlsx'):
        df.to_excel(file_path, index=False)
    else:
        raise ValueError("Unsupported file format. Please use CSV or Excel (.xlsx).") 