from sqlalchemy import create_engine, text
from app_tkinter import Base

def migrate_database():
    """Add cost_center column to employees table if it doesn't exist"""
    engine = create_engine('sqlite:///forecast_tool.db')
    
    with engine.connect() as conn:
        # Check if cost_center column exists
        result = conn.execute(text("SELECT name FROM pragma_table_info('employees') WHERE name='cost_center'"))
        if not result.fetchone():
            # Add cost_center column
            conn.execute(text("ALTER TABLE employees ADD COLUMN cost_center VARCHAR"))
            conn.execute(text("UPDATE employees SET cost_center = manager_code"))  # Set default value
            print("Added cost_center column to employees table")
        
        conn.commit()

if __name__ == '__main__':
    migrate_database() 