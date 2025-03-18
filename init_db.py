from models import init_db, get_session, Employee, GA01Week, WorkCode, Forecast
from datetime import date
import datetime

def initialize_database():
    # Create database and tables
    init_db()
    session = get_session()
    
    # Initialize work codes
    work_codes = [
        {"code": "Work Time", "description": "Regular work time"},
        {"code": "Project Time", "description": "Time spent on projects"}
    ]
    
    for wc_data in work_codes:
        wc = WorkCode(
            code=wc_data["code"],
            description=wc_data["description"]
        )
        session.add(wc)
    
    # Initialize default GA01 weeks for current year
    current_year = datetime.datetime.now().year
    default_weeks = {
        1: 4.0,  # January
        2: 4.0,  # February
        3: 4.33, # March
        4: 4.0,  # April
        5: 4.33, # May
        6: 4.33, # June
        7: 4.0,  # July
        8: 4.33, # August
        9: 4.0,  # September
        10: 4.33, # October
        11: 4.33, # November
        12: 4.0,  # December
    }
    
    for month, weeks in default_weeks.items():
        ga01 = GA01Week(
            year=current_year,
            month=month,
            weeks=weeks
        )
        session.add(ga01)
    
    session.commit()
    session.close()
    
    print("Database initialized with default work codes and GA01 weeks")

if __name__ == "__main__":
    initialize_database() 