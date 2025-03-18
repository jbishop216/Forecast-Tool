import tkinter as tk
from tkinter import ttk, messagebox

class ForecastApp(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Employee Forecasting Tool")
        self.geometry("800x500")
        
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Employee Forecasting Tool", font=("Helvetica", 16)).pack(pady=20)
        
        ttk.Label(main_frame, text="Instructions:").pack(anchor=tk.W, pady=5)
        ttk.Label(main_frame, text="1. This application requires additional dependencies:").pack(anchor=tk.W)
        ttk.Label(main_frame, text="   - pandas, SQLAlchemy, openpyxl").pack(anchor=tk.W)
        ttk.Label(main_frame, text="2. Please install them using:").pack(anchor=tk.W, pady=5)
        ttk.Label(main_frame, text="   pip install pandas SQLAlchemy openpyxl").pack(anchor=tk.W)
        ttk.Label(main_frame, text="3. Then run the full application using:").pack(anchor=tk.W, pady=5) 
        ttk.Label(main_frame, text="   python main.py").pack(anchor=tk.W)
        
        ttk.Separator(main_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=20)
        
        # Test button to show everything is working
        ttk.Button(main_frame, text="Click to Test Tkinter", 
                   command=lambda: messagebox.showinfo("Success", "Tkinter is working correctly!")).pack(pady=10)
        
        exit_button = ttk.Button(main_frame, text="Exit", command=self.destroy)
        exit_button.pack(pady=10)

if __name__ == "__main__":
    app = ForecastApp()
    app.mainloop() 