from models import init_db, get_session, Employee, GA01Week, WorkCode, Settings, ProjectAllocation, PlannedChange
from datetime import datetime

def initialize_database():
    """Initialize the database with required tables and default data"""
    # Create tables
    init_db()
    
    session = get_session()
    
    # Check if work codes exist
    work_codes = session.query(WorkCode).all()
    if not work_codes:
        # Add default work codes
        work_time = WorkCode(code="Work Time", description="Regular work time")
        project_time = WorkCode(code="Project Time", description="Project allocation time")
        
        session.add(work_time)
        session.add(project_time)
        session.commit()
    
    # Initialize settings if they don't exist
    settings = session.query(Settings).first()
    if not settings:
        settings = Settings(fte_hours=34.5, contractor_hours=39.0)
        session.add(settings)
        session.commit()
    
    # Initialize GA01 weeks for current year if they don't exist
    current_year = datetime.now().year
    ga01_weeks = session.query(GA01Week).filter(GA01Week.year == current_year).all()
    
    if not ga01_weeks:
        # Add default GA01 weeks (4.0 weeks per month)
        for month in range(1, 13):
            ga01 = GA01Week(year=current_year, month=month, weeks=4.0)
            session.add(ga01)
        session.commit()
    
    session.close()
    
    print("Database initialized successfully!")

if __name__ == "__main__":
    initialize_database() 