# Forecast Tool

A desktop application for managing employee project allocations and forecasting hours.

## Features

- **Employee Management**: Add, edit, and remove employees with their details
- **GA01 Week Configuration**: Set up GA01 weeks for accurate time calculations
- **Forecast Calculation**: Calculate and export monthly forecasts based on project allocations
- **Project Time Allocation**: Manage project time allocations by manager code and cost center
- **Planned Employee Changes**: Schedule and manage future employee changes
  - New hires
  - Terminations
  - Employment type conversions
- **Visualizations**: Generate charts and graphs for:
  - Monthly forecast hours
  - Manager allocations
  - Employee type distribution
- **Settings**: Configure FTE and contractor weekly hours
- **Excel Integration**: Export forecast data to Excel

## Tabs

1. **Employees**: Manage employee information (name, manager code, employment type, start/end dates)
2. **GA01 Weeks**: Set the number of working weeks for each month of the year
3. **Forecast**: Calculate and view forecasts based on project allocations and GA01 weeks
4. **Project Allocation**: Manage project allocations (cost center, work code, monthly hours)
5. **Planned Changes**: Track employee changes (new hires, conversions, terminations)
6. **Visualizations**: View various charts and analytics

## Technical Details

- Built with Python 3.11 and Tkinter
- SQLite database for data storage
- SQLAlchemy ORM for database interaction
- Matplotlib for data visualization
- OpenPyXL for Excel export

## Requirements

```txt
sqlalchemy
matplotlib
openpyxl
```

## Installation

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd forecast-tool
   ```

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Initialize the database:
   ```bash
   python app_tkinter.py
   ```

## Usage

1. Start by adding employees in the Employees tab
2. Configure GA01 weeks for accurate time calculations
3. Add project allocations for each manager
4. Use the Forecast tab to calculate and view forecasts
5. Export data to Excel as needed
6. Track planned changes in the Planned Changes tab
7. View analytics in the Visualizations tab

## Mid-Month Employee Changes

The tool handles mid-month employee changes based on GA01 weeks:
- If changes (hiring, termination, conversion) occur after the 2nd GA01 week, employees are counted for the entire month
- If changes occur before the end of the 2nd GA01 week, they are prorated accordingly

## Settings

Access the Settings dialog from the File menu to configure:
- FTE weekly hours (default: 34.5)
- Contractor weekly hours (default: 39.0) 