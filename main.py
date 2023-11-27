import math
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk
import calendar
from datetime import datetime
from collections import Counter

path = os.path.dirname(os.path.abspath(__file__))
print("PATH: ", path)


def total_working_hours(path):
    data = []
    dates = []
    duration = []
    counter = 0
    total_hours = 0
    total_minutes = 0
    remaining_minutes = 0
    dir = os.listdir(path)

    for file in dir:
        data.append(file)
    
    for file in data:
        if os.path.splitext(file)[1].lstrip(".") == "xlsx":
            excel_sheet = pd.read_excel(file)
            excel_sheet.dropna()
        
            hours = excel_sheet['total hours'][0].hour
            minutes = excel_sheet['total hours'][0].minute
            dates.append(f"{int(excel_sheet['year'][0])}-{int(excel_sheet['month'][0]):02d}-{int(excel_sheet['day'][0]):02d}")
            duration.append(f"{hours}h {minutes}m")   
       
            total_hours += hours
            total_minutes += minutes

    if total_minutes > 59:
        total_hours += (total_minutes // 60)
        remaining_minutes += (total_minutes % 60)
    else:
        remaining_minutes = total_minutes

    return dates, total_hours, duration

def GUI():

    dates, total_hours, duration = total_working_hours(path)
    # rounding to the nearest multiple of 10
    total_hours = int(math.ceil(total_hours/ 10.0)) * 10

    root = tk.Tk()
    root.title("Working hours")
    root.geometry("800x700")

    frame = ttk.Frame(root)
    frame.pack(fill="both", expand=True)
    table = ttk.Treeview(frame, columns=("Hours", "Minutes"))

    table.heading("#1", text="Dates")
    table.heading("#2", text="Working hours")

    
    for i in range(len(duration)):
        table.insert("", "end", text=str(i), values=(dates[i], duration[i]))
       
    last_row = table.insert("", "end", text="Total hours", values=("", f"{total_hours}h")) 
    table.item(last_row, tags=("last_row",)) 
    table.tag_configure("last_row", background="light green")   
    
  
    table.pack(fill="both", expand=True)

    vsb = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
    table.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")    

    root.mainloop()

# only work days 
# every day less than 8
# no two files for the same date
# cf. the general internship guide, e.g. take a rest after 6 hours
# Round the final hours to 10 >>> so 198 will be 200

# for checking if day worked was a working day
def is_working_day(dates, holidays):
    flag = False
    for date in dates:
        date_str_to_object = datetime.strptime(date, "%Y-%m-%d")
        if date_str_to_object.weekday() <= 0 and date_str_to_object.weekday() >= 4 and date not in holidays:
            print(f"{date} is not a working day")
            flag = True
    if not flag:
        print("All files have working hours on a working day")

# checks if working hours is less than 8   
def is_working_hours_valid(dates, duration): 
    hours_list = [int(time.split('h')[0]) for time in duration]
    flag = False
    
    for i in range(len(hours_list)):
        if hours_list[i] > 8:
            print(f"{dates[i]} has invalid working hours (> 8)")
            flag = True
    if not flag:
        print("All files have valid working hours")
            
# checks for file duplicates   
def is_file_valid():
    file_names = [file for file in os.listdir(path) if os.path.splitext(file)[1].lstrip(".") == "xlsx"]
    file_counts = Counter(file_names)
    flag = False
    
    for file_name, count in file_counts.items():
        if count > 1:
            print(f"This file: {file_name} is a duplicate")
            flag = True
    if not flag:    
        print("All files are valid")
    
# checks whether there are duplicate dates within files    
def check_dates(dates):
    flag = False
    reported_pairs = set()  
    for i in range(len(dates)):
        for j in range(i + 1, len(dates)):  
            if dates[i] == dates[j] and (i, j) not in reported_pairs and (j, i) not in reported_pairs:
                print(f"There are duplicate dates files: {dates[i]} and {dates[j]}")
                flag = True
                reported_pairs.add((i, j))  
    if not flag:
        print("No files have duplicate dates")


def validate():
    dates, total_hours, duration = total_working_hours(path)
    holidays = ['2023-09-20', '2023-10-03', '2023-10-31', '2023-11-01', '2023-11-22']
    
    is_working_day(dates, holidays)
    is_working_hours_valid(dates, duration)
    check_dates(dates)
    is_file_valid()
    
def main():
    GUI()
    validate()
    
main()  