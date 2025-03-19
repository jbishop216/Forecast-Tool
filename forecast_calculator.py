from models import Employee, GA01Week, WorkCode, Forecast, ProjectAllocation, PlannedChange, ChangeType, get_session
from datetime import date
import calendar

def calculate_employee_forecast(employee, year, work_code_id):
    """Calculate monthly forecast hours for an employee for a specific year and work code"""
    forecasts = []
    session = get_session()
    
    # Get GA01 weeks for the year
    ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
    ga01_dict = {week.month: week.weeks for week in ga01_weeks}
    
    # Get work code
    work_code = session.query(WorkCode).filter(WorkCode.id == work_code_id).first()
    
    # Get planned changes that affect this employee
    planned_changes = []
    
    # Get termination or conversion changes
    direct_changes = session.query(PlannedChange).filter(
        PlannedChange.employee_id == employee.id,
        PlannedChange.effective_date.between(date(year, 1, 1), date(year, 12, 31))
    ).order_by(PlannedChange.effective_date).all()
    
    planned_changes.extend(direct_changes)
    
    # For each month of the year
    for month in range(1, 13):
        # If GA01 data not available for this month, skip
        if month not in ga01_dict:
            continue
            
        # Get first and last day of month
        last_day = calendar.monthrange(year, month)[1]
        month_start = date(year, month, 1)
        month_end = date(year, month, last_day)
        
        # Calculate the day after the end of the 2nd GA01 week (if there are at least 2 weeks)
        weeks_in_month = ga01_dict[month]
        second_week_threshold = None
        if weeks_in_month >= 2:
            # Calculate the day that represents the end of the 2nd week
            # Assuming each week is 7 days, the 2nd week would end on day 14
            second_week_threshold = date(year, month, min(14, last_day))
        
        # Apply any planned changes that would affect this month
        this_month_changes = []
        for change in planned_changes:
            if month_start <= change.effective_date <= month_end:
                this_month_changes.append(change)
        
        # Skip this month if employee is terminated
        has_termination = False
        for change in this_month_changes:
            if change.change_type == ChangeType.TERMINATION.value:
                # Check if termination is after 2nd week threshold
                if second_week_threshold and change.effective_date > second_week_threshold:
                    # Count them for the entire month
                    pass
                else:
                    # Terminate according to actual date
                    has_termination = True
                    break
                    
        if has_termination:
            hours = 0
        # Handle the case when employee hasn't started yet or has already ended
        elif employee.start_date and employee.start_date > month_end:
            # Future start date beyond this month
            hours = 0
        elif employee.end_date and employee.end_date < month_start:
            # Already ended before this month
            hours = 0
        else:
            # Calculate days in month the employee is active, with special handling for mid-month changes
            
            # Handle start date considering the 2nd GA01 week rule
            if employee.start_date and month_start <= employee.start_date <= month_end:
                # If they start after the 2nd week threshold, count them for the whole month
                if second_week_threshold and employee.start_date <= second_week_threshold:
                    start_day = employee.start_date
                else:
                    start_day = month_start  # Count them for the entire month
            else:
                start_day = month_start
                
            # Handle end date considering the 2nd GA01 week rule
            if employee.end_date and month_start <= employee.end_date <= month_end:
                # If they end after the 2nd week threshold, count them for the whole month
                if second_week_threshold and employee.end_date > second_week_threshold:
                    end_day = month_end  # Count them for the entire month
                else:
                    end_day = employee.end_date
            else:
                end_day = month_end
            
            # Handle employee type changes (conversions) during month
            employment_type = employee.employment_type
            conversion_changes = [c for c in this_month_changes if c.change_type == ChangeType.CONVERSION.value]
            
            if conversion_changes:
                # Use the last planned conversion in the month
                latest_conversion = conversion_changes[-1]
                employment_type = latest_conversion.target_employment_type
            
            # Calculate active ratio based on dates
            active_days = (end_day - start_day).days + 1
            total_days = (month_end - month_start).days + 1
            active_ratio = active_days / total_days
            
            # Get weekly hours based on employment type
            weekly_hours = 34.5 if employment_type == "FTE" else 39.0
            
            # Calculate total available hours
            total_hours = weekly_hours * ga01_dict[month] * active_ratio
            
            # Get project allocation for this manager code and month
            project_allocation = session.query(ProjectAllocation).filter(
                ProjectAllocation.manager_code == employee.manager_code,
                ProjectAllocation.year == year,
                ProjectAllocation.month == month
            ).first()
            
            # Determine hours based on work code type
            if work_code.code == "Project Time" and project_allocation:
                # For project time, use the manually entered allocation
                hours = project_allocation.hours * active_ratio
            elif work_code.code == "Work Time":
                # For work time, subtract project hours from total available hours
                project_hours = 0
                if project_allocation:
                    project_hours = project_allocation.hours * active_ratio
                hours = max(0, total_hours - project_hours)
            else:
                # For any other code, use previous 80/20 split as fallback
                if work_code.code == "Project Time":
                    hours = total_hours * 0.2
                else:
                    hours = total_hours * 0.8
        
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
        
        # Also include planned new hires that haven't been applied yet
        new_hire_plans = session.query(PlannedChange).filter(
            PlannedChange.change_type == ChangeType.NEW_HIRE.value,
            PlannedChange.effective_date.between(date(current_year, 1, 1), date(current_year + 1, 12, 31))  # Include next year too
        ).all()
        
        for plan in new_hire_plans:
            # Create a temporary employee object for the planned hire
            new_hire = Employee(
                id=-plan.id,  # Use negative plan ID to avoid conflicts
                name=plan.name,
                team=plan.team,
                manager_code=plan.manager_code,
                cost_center=plan.cost_center,
                start_date=plan.effective_date,
                employment_type=plan.employment_type
            )
            employees.append(new_hire)
    
    # Delete existing forecasts for these employees
    for emp in employees:
        # Skip temporary employees (planned hires)
        if emp.id > 0:
            session.query(Forecast).filter(Forecast.employee_id == emp.id).delete()
    
    # Delete any forecasts for planned hires
    session.query(Forecast).filter(Forecast.employee_id.is_(None)).delete()
    
    # Calculate new forecasts
    for emp in employees:
        for work_code in work_codes:
            forecasts = calculate_employee_forecast(emp, current_year, work_code.id)
            for f_data in forecasts:
                # For planned hires
                if emp.id < 0:
                    # Create a "future hire" note in the employee name
                    f_data["employee_id"] = None
                    f_data["notes"] = f"Future hire: {emp.name} ({emp.manager_code}, {emp.cost_center}) planned {emp.start_date}"
                
                forecast = Forecast(**f_data)
                session.add(forecast)
    
    session.commit()
    session.close() 