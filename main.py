import os
import pandas as pd
import tkinter as tk
from tkinter import ttk


path = os.path.dirname(os.path.abspath(__file__))


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

    return dates, total_hours, remaining_minutes, duration

def GUI():

    dates, total_hours, remaining_minutes, duration = total_working_hours(path)

    root = tk.Tk()
    root.title("Working hours")
    root.geometry("800x700")

    frame = ttk.Frame(root)
    frame.pack(fill="both", expand=True)
    table = ttk.Treeview(frame, columns=("Hours", "Minutes"))

    table.heading("#1", text="Dates")
    table.heading("#2", text="Duration")

    
    for i in range(len(duration)):
        table.insert("", "end", text=str(i), values=(dates[i], duration[i]))
       
    last_row = table.insert("", "end", text="Total hours", values=("", f"{total_hours}h {remaining_minutes}m")) 
    table.item(last_row, tags=("last_row",)) 
    table.tag_configure("last_row", background="light green")   
    
  
    table.pack(fill="both", expand=True)

    vsb = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
    table.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")    

    root.mainloop()


def main():
    GUI()

main()