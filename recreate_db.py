import os
from models import init_db, Base
from sqlalchemy import create_engine

def recreate_database():
    """Drop and recreate the database"""
    # Delete the existing database file
    if os.path.exists('forecast_tool.db'):
        os.remove('forecast_tool.db')
        print("Existing database removed.")
    
    # Create a new database
    engine = init_db()
    print("New database created.")
    
    # Run the initialization script
    from init_db import initialize_database
    initialize_database()

if __name__ == "__main__":
    recreate_database() 