import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt

Vorname_list = [
    "Adelina", "Johannes", "Lucia", "Anne", "Frank", "Fabian", "Robin",
    "Raphael", "Silke", "Kerstin", "Arno", "Marcus", "Alfons", "Nils",
    "Tim", "Achim", "Igor", "Lukas"
]

def main(file_path, month=None, employee_name=None, start_time=None, end_time=None):
    global Nr_Zeichner

    df_origin = pd.read_excel(file_path, sheet_name="Kapazität intern 2025", header=1)
    df_origin = pd.DataFrame(df_origin)

    df = df_origin.drop(df_origin.index[-7:], inplace=False)
    df = df.drop(df_origin.index[:4])
    df = df.drop(columns=['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 371'])

    df.rename(columns={'heute': 'Name', 'Unnamed: 2': 'Vorname'}, inplace=True)

    df.columns = df.columns.astype(str)
    df.columns = [col.replace(' 00:00:00', '') if i >= 2 and i < 367 else col for i, col in enumerate(df.columns)]
    df = df.fillna('X')

    df = df[df['Vorname'].isin(Vorname_list)]

    df.iloc[:, 2:] = df.iloc[:, 2:].astype(str)
    df.reset_index(drop=True, inplace=True)

    Nr_Zeichner = len(df)

    if month:
        workdays_monthly(df, month)
    if employee_name:
        employee_working_days(df, employee_name)
    if start_time and end_time:
        check_given_time(df, start_time, end_time)
    print("Plot finished")

def workdays_monthly(df, month):
    month_columns = [col for col in df.columns[2:] if col.startswith(f"2025-{month:02d}")]
    
    df['Working Days'] = df.loc[0:Nr_Zeichner+1, month_columns].apply(lambda row: row.isin(['h', '1']).sum(), axis=1)
    
    plt.figure(figsize=(10, 6))
    bars = plt.bar(df.loc[0:Nr_Zeichner, 'Vorname'], df.loc[0:Nr_Zeichner, 'Working Days'], color='skyblue')
    plt.xlabel('Employee')
    plt.ylabel('Number of Working Days')
    plt.title(f'Number of Working Days in {month:02d}/2025')
    plt.xticks(rotation=45)
    
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom')
    
    plt.tight_layout()
    plt.show(block=False)

def employee_working_days(df, employee_name):
    employee_row = df[df['Vorname'] == employee_name]
    
    working_days_per_month = {}
    
    for month in range(1, 13):
        month_columns = [col for col in df.columns[2:] if col.startswith(f"2025-{month:02d}")]
        
        working_days = employee_row[month_columns].apply(lambda row: row.isin(['h', '1']).sum(), axis=1).values[0]
        
        working_days_per_month[f"2025-{month:02d}"] = working_days
    
    plt.figure(figsize=(10, 6))
    plt.bar(working_days_per_month.keys(), working_days_per_month.values(), color='skyblue')
    plt.xlabel('Month')
    plt.ylabel('Number of Working Days')
    plt.title(f'{employee_name} in 2025')
    plt.xticks(rotation=45)
    
    for month, days in working_days_per_month.items():
        plt.text(month, days + 0.1, int(days), ha='center', va='bottom')
    
    plt.tight_layout()
    plt.show(block=False)

def check_given_time(df, start_time, end_time):
    time_columns = [col for col in df.columns[2:] if start_time <= col <= end_time]

    df['Working Days'] = df.loc[0:Nr_Zeichner+1, time_columns].apply(lambda row: row.isin(['h', '1']).sum(), axis=1)

    plt.figure(figsize=(10, 6))
    bars = plt.bar(df.loc[0:Nr_Zeichner+1, 'Vorname'], df.loc[0:Nr_Zeichner+1, 'Working Days'], color='skyblue')
    plt.xlabel('Employee')
    plt.ylabel('Number of Working Days')
    plt.title(f'Number of Working Days from {start_time} to {end_time}')
    plt.xticks(rotation=45)
    
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom')
    
    plt.tight_layout()
    plt.show()

def browse_file():
    file_path = filedialog.askopenfilename()
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

def run_main_month():
    file_path = file_entry.get()
    month = int(month_entry.get())
    main(file_path, month=month)

def run_main_employee():
    file_path = file_entry.get()
    employee_name = employee_entry.get()
    main(file_path, employee_name=employee_name)

def run_main_time():
    file_path = file_entry.get()
    start_time = start_time_entry.get()
    end_time = end_time_entry.get()
    main(file_path, start_time=start_time, end_time=end_time)

root = tk.Tk()
root.title("Mini APP - Personalplanung Visualisierung")

notebook = ttk.Notebook(root)

# Tab 1: File Path
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="File Path")

tk.Label(tab1, text="Excel File:").grid(row=0, column=0, padx=10, pady=10)
file_entry = tk.Entry(tab1, width=50)
file_entry.grid(row=0, column=1, padx=10, pady=10)
browse_button = tk.Button(tab1, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# Tab 2: Month
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Month")

tk.Label(tab2, text="Month (1-12):").grid(row=0, column=0, padx=10, pady=10)
month_entry = tk.Entry(tab2, width=10)
month_entry.grid(row=0, column=1, padx=10, pady=10)

run_button_month = tk.Button(tab2, text="Run", command=run_main_month)
run_button_month.grid(row=1, column=0, columnspan=2, pady=20)

# Tab 3: Employee Name
tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="Employee Name")

tk.Label(tab3, text="Employee Name:").grid(row=0, column=0, padx=10, pady=10)
employee_entry = tk.Entry(tab3, width=20)
employee_entry.grid(row=0, column=1, padx=10, pady=10)

run_button_employee = tk.Button(tab3, text="Run", command=run_main_employee)
run_button_employee.grid(row=1, column=0, columnspan=2, pady=20)

# Tab 4: Time Range
tab4 = ttk.Frame(notebook)
notebook.add(tab4, text="Time Range")

tk.Label(tab4, text="Start Time (YYYY-MM-DD):").grid(row=0, column=0, padx=10, pady=10)
start_time_entry = tk.Entry(tab4, width=20)
start_time_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(tab4, text="End Time (YYYY-MM-DD):").grid(row=1, column=0, padx=10, pady=10)
end_time_entry = tk.Entry(tab4, width=20)
end_time_entry.grid(row=1, column=1, padx=10, pady=10)

run_button_time = tk.Button(tab4, text="Run", command=run_main_time)
run_button_time.grid(row=2, column=0, columnspan=2, pady=20)

notebook.pack(expand=True, fill='both')

root.mainloop()