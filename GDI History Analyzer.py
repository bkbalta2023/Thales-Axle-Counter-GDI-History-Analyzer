#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import re
import os
import subprocess
from PIL import Image, ImageTk
import time

excel_file_path = ""

def open_files():
    global excel_file_path
    htm_filenames = filedialog.askopenfilenames(title="Select .htm Files", filetypes=[("HTML Files", "*.htm")])
    if not htm_filenames:
        return
    
    analyzing_label.pack()
    root.update()
    
    time.sleep(1.5)  # Display "Analyzing..." for 1.5 seconds
    
    excel_save_path = os.path.join(os.path.dirname(htm_filenames[0]), "GDI History Analyze.xlsx")
    excel_file_path = excel_save_path
    
    with pd.ExcelWriter(excel_save_path, engine='xlsxwriter') as writer:
        for htm_filename in htm_filenames:
            with open(htm_filename, 'r') as file:
                data = file.readlines()

            parsed_data = []
            pattern = r'(\d+/\d+/\d+ \d+:\d+:\d+): (Detection Point \d+); Detection-Point-Information: (.+)'
            for line in data:
                match = re.match(pattern, line)
                if match:
                    datetime_str = match.group(1)
                    point = match.group(2)
                    info = match.group(3).replace('<br>', '')
                    
                    if "no more" in info.lower() or "5 minutes" in info.lower():
                        continue
                    
                    parsed_data.append([datetime_str, point, info])

            if parsed_data:
                columns = ["Datetime", "Detection Point", "Detection-Point-Information"]
                df = pd.DataFrame(parsed_data, columns=columns)
                sheet_name = os.path.basename(htm_filename).replace('.htm', '')
                
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                worksheet = writer.sheets[sheet_name]
                for idx, col in enumerate(df):
                    series = df[col]
                    max_len = max((
                        series.astype(str).map(len).max(),
                        len(str(series.name))
                    )) + 1
                    worksheet.set_column(idx, idx, max_len)
                worksheet.autofilter(0, 0, df.shape[0], df.shape[1] - 1)

        result_label.config(text=f"Data saved to: {excel_save_path}")
        open_excel_button.config(state=tk.NORMAL)
        
        time.sleep(2)  # Wait for 2 seconds after the data is saved
        
        analyzing_label.pack_forget()

def open_excel():
    global excel_file_path
    if excel_file_path:
        subprocess.Popen(["start", "excel", excel_file_path], shell=True)
    else:
        result_label.config(text="No Excel file to open.")

root = tk.Tk()
root.title("GDI History Analyzer")
root.geometry("800x600")

style = ttk.Style()
style.theme_use("clam")

style.configure("TButton", font=("Helvetica", 12), foreground="white", background="#007acc")
style.configure("TLabel", font=("Helvetica", 14), foreground="#333", background="#f0f0f0")
style.configure("TNotebook", background="#f0f0f0")
style.configure("blue.Horizontal.TProgressbar", background="blue")

notebook = ttk.Notebook(root)
notebook.pack(pady=20, fill="both", expand=True)

tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Import .htm Files")  # Change the tab text here

open_button = ttk.Button(tab1, text="Import .htm Files", command=open_files)  # Change the button text here
open_button.pack(pady=20)

open_excel_button = ttk.Button(tab1, text="Open Excel File", command=open_excel, state=tk.DISABLED)
open_excel_button.pack(pady=10)

analyzing_label = tk.Label(tab1, text="Analyzing...", font=("Helvetica", 12))

result_label = tk.Label(tab1, text="", font=("Helvetica", 12), fg="green")
result_label.pack()

developer_label = tk.Label(root, text="Developed by Burak Kagan Balta", anchor="se", bg="#f0f0f0", fg="#666", font=("Helvetica", 10))
developer_label.pack(side="bottom", padx=10, pady=10)

root.mainloop()

