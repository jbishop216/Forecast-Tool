import sys
import os
import tkinter.messagebox as messagebox
from app_tkinter import ForecastApp
from models import init_db
from init_db import initialize_database

if __name__ == "__main__":
    # Create database if it doesn't exist
    db_file = "forecast_tool.db"
    if not os.path.exists(db_file):
        try:
            initialize_database()
            print("Database initialized successfully")
        except Exception as e:
            print(f"Error initializing database: {e}")
            messagebox.showerror("Database Error", f"Failed to initialize database: {str(e)}")
            sys.exit(1)
    
    # Start the application
    app = ForecastApp()
    app.mainloop() 