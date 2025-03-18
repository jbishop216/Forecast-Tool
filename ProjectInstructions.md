Project Name: Employee Forecasting Tool

Objective:
Develop a GUI-based application that forecasts the number of hours employees will charge per month for the remainder of each year. The tool should consider employee type (FTE or Contractor), manager code, and cost center while dynamically adjusting to employee changes throughout the year.

Core Requirements:
	1.	User Interface (GUI)
	•	A structured tabular interface displaying the entire year (12-month forecast).
	•	Rows represent work codes (e.g., “Work Time” and “Project Time”) by Manager Code and Cost Center.
	•	Columns represent months, each with a GA01 week multiplier (manually adjustable per year).
	•	Dynamic editing of employees, work codes, and GA01 weeks with real-time updates to forecasts.
	•	Export functionality to CSV or Excel for reporting.
	2.	Data Structure & Storage
	•	Persistent storage (SQLite recommended) to retain data across sessions.
	•	Tables required:
	•	Employees: Tracks name, team, manager code, cost center, start date, end date, and employment type (FTE or Contractor).
	•	GA01 Weeks: Stores the weeks per month used to calculate forecasted hours.
	•	Forecasts: Stores computed forecasted hours based on employees and GA01 weeks.
	3.	Forecast Calculation Rules:
	•	FTEs work 34.5 hours per week; Contractors work 39 hours per week.
	•	Monthly hours = weekly hours × GA01 weeks for that month.
	•	Forecast adjusts automatically if an employee changes employment type or leaves mid-year.
	•	Each work code (Work Time, Project Time) is calculated separately.
	4.	Import & Export Features:
	•	Ability to import employee lists and GA01 weeks from CSV/Excel.
	•	Ability to export the full-year forecast for review.

Technical Considerations:
	•	Python GUI: Tkinter or PyQt (whichever offers better UI flexibility).
	•	Database: SQLite for persistence and scalability.
	•	Pandas & OpenPyXL: For handling CSV/Excel imports/exports.

Development Roadmap:
	1.	Set up SQLite database & schema for Employees, GA01 Weeks, and Forecasts.
	2.	Build the core UI (table-based view for forecasts, editable fields for GA01 weeks and employees).
	3.	Implement forecasting logic (adjusts based on employee type and mid-year changes).
	4.	Enable import/export features for CSV/Excel compatibility.
	5.	Testing & validation to ensure accurate calculations and a user-friendly experience.

Final Deliverable:

A fully functional, persistent, and interactive forecasting tool that allows manual adjustments, real-time recalculations, and seamless import/export for planning employee work hours.