from models import Employee, GA01Week, WorkCode, Forecast, get_session
from datetime import date
import calendar

def calculate_employee_forecast(employee, year, work_code_id):
    """Calculate monthly forecast hours for an employee for a specific year and work code"""
    forecasts = []
    session = get_session()
    
    # Get GA01 weeks for the year
    ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
    ga01_dict = {week.month: week.weeks for week in ga01_weeks}
    
    # For each month of the year
    for month in range(1, 13):
        # If GA01 data not available for this month, skip
        if month not in ga01_dict:
            continue
            
        # Get first and last day of month
        last_day = calendar.monthrange(year, month)[1]
        month_start = date(year, month, 1)
        month_end = date(year, month, last_day)
        
        # If employee hasn't started yet or has already ended before this month
        if (employee.start_date and employee.start_date > month_end) or \
           (employee.end_date and employee.end_date < month_start):
            hours = 0
        else:
            # Calculate days in month the employee is active
            start_day = max(month_start, employee.start_date) if employee.start_date else month_start
            end_day = min(month_end, employee.end_date) if employee.end_date else month_end
            
            active_days = (end_day - start_day).days + 1
            total_days = (month_end - month_start).days + 1
            
            # Proportion of month the employee is active
            active_ratio = active_days / total_days
            
            # Calculate hours based on weekly hours, GA01 weeks, and active ratio
            hours = employee.weekly_hours * ga01_dict[month] * active_ratio
            
            # For simplicity, assume "Work Time" gets 80% of hours and "Project Time" gets 20%
            # This could be made more sophisticated based on additional requirements
            work_code = session.query(WorkCode).filter(WorkCode.id == work_code_id).first()
            if work_code.code == "Work Time":
                hours *= 0.8
            elif work_code.code == "Project Time":
                hours *= 0.2
        
        # Create Forecast object
        forecast = {
            "employee_id": employee.id,
            "work_code_id": work_code_id,
            "year": year,
            "month": month,
            "hours": round(hours, 2)
        }
        
        forecasts.append(forecast)
    
    session.close()
    return forecasts

def update_forecast(employee_id=None):
    """Update forecasts for one employee or all employees"""
    session = get_session()
    current_year = date.today().year
    
    # Get work codes
    work_codes = session.query(WorkCode).all()
    
    if employee_id:
        employees = [session.query(Employee).filter(Employee.id == employee_id).first()]
    else:
        employees = session.query(Employee).all()
    
    # Delete existing forecasts for these employees
    for emp in employees:
        session.query(Forecast).filter(Forecast.employee_id == emp.id).delete()
    
    # Calculate new forecasts
    for emp in employees:
        for work_code in work_codes:
            forecasts = calculate_employee_forecast(emp, current_year, work_code.id)
            for f_data in forecasts:
                forecast = Forecast(**f_data)
                session.add(forecast)
    
    session.commit()
    session.close() 