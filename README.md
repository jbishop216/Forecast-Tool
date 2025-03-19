# Forecast Tool

A desktop application for managing employee project allocations and forecasting hours.

## Features

- **Employee Management**: Add, edit, and remove employees with their details
- **Project Time Allocation**: Manage project time allocations by manager code
- **GA01 Week Configuration**: Set up GA01 weeks for accurate time calculations
- **Forecast Calculation**: Calculate monthly forecasts based on employee data
- **Planned Employee Changes**: Schedule and manage future employee changes
  - New hires
  - Terminations
  - Employment type conversions
- **Excel Integration**: Copy and paste data to and from Excel

## Mid-Month Employee Changes

The tool handles mid-month employee changes based on GA01 weeks:
- If changes (hiring, termination, conversion) occur after the 2nd GA01 week, employees are counted for the entire month
- If changes occur before the end of the 2nd GA01 week, they are prorated accordingly

## Technical Details

- Built with Python and Tkinter
- SQLite database for data storage
- SQLAlchemy ORM for database interaction

## Installation

1. Clone this repository
2. Run `python init_db.py` to initialize the database
3. Run `python app_tkinter.py` to start the application 