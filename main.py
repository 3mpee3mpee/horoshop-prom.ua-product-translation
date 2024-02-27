import pandas as pd
import concurrent.futures
from tkinter import Tk, Button, Label, filedialog, ttk
from mtranslate import translate
import time
import webbrowser
import traceback
import sys

# Translation Functions
def translate_text(text, heading):
    try:
        if isinstance(text, str):
            text_with_heading = f"{heading}: {text}"
            translated_text = translate(text_with_heading, 'uk')
            formatted_text = translated_text[translated_text.find(': ') + 2:]
            return formatted_text
    except Exception as e:
        print(e)
        print(traceback.format_exc())
        if "HTTP Error 429: Too Many Requests" in str(e): 
            show_alert("Rate Limit Exceeded: please try again in 1 hr")
            sys.exit(0)
        time.sleep(8)
        return translate_text(text, heading)

def translate_excel(input_file, output_file, status_label):
    start_time = time.time()
    status_label.config(text="Executing...")

    # Load the file
    df = pd.read_excel(input_file)
    header_rows = df.columns.ravel()

    total_progress = len(header_rows[2:]) * len(df)  
    current_progress = 0  

    # Process each heading
    for heading in header_rows[1:]:  
        if "(RU)" not in heading and "(UA)" not in heading:
            new_heading_ru = heading + "(RU)"
            new_heading_ua = heading + "(UA)"
            new_row_ru = df[heading].copy()
            new_row_ua = df[heading].copy()
            df[new_heading_ru] = new_row_ru
            df[new_heading_ua] = new_row_ua

    header_rows = df.columns.ravel()
    for heading in header_rows[2:]:
        if "(UA)" in heading:
            for i in range(len(df)):
                text = df.loc[i, heading]
                translated_text = translate_text(text, heading)
                df.loc[i, heading] = translated_text
                current_progress += 1
                progress_percent = min(100, int((current_progress / total_progress) * 100))
                root.after(10, lambda: status_label.config(text=f"Executing... {progress_percent}%"))

    # Save the result
    df.to_excel(output_file)
    end_time = time.time()
    execution_time = end_time - start_time
    status_label.config(text=f"Finished. Execution time: {execution_time:.2f} seconds")

# GUI Functions
def show_alert(message):
    alert_window = Tk()
    alert_window.title("Alert")
    alert_label = Label(alert_window, text=message)
    alert_label.pack(padx=20, pady=10)
    ok_button = Button(alert_window, text="OK", command=alert_window.destroy)
    ok_button.pack(pady=10)
    alert_window.mainloop()

def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if file_paths:
        process_files(file_paths)

def process_files(file_paths):
    for file_path in file_paths:
        output_file = file_path.replace(".xlsx", "_translated.xlsx")
        Label(root, text=file_path).pack()
        status_label = Label(root, text="")
        status_label.pack()
        executor.submit(translate_excel, file_path, output_file, status_label)

def callback(event):
    webbrowser.open_new("https://github.com/3mpee3mpee")

def on_closing():
    root.destroy()

# GUI Setup
root = Tk()
root.title("Excel Translator")
root.geometry("500x500")

title_label = Label(root, text="Excel Translator", font=("Helvetica", 16))
title_label.pack(pady=10)

description_label = Label(root, text="Translate specifications from Prom.ua to Horoshop's format.", wraplength=400)
description_label.pack(pady=5)

author_label = ttk.Label(root, text="made by @3mpee3mpee", font=("Helvetica", 10), style="Italic.TLabel")
author_label.pack(pady=5)
author_label.bind("<Button-1>", callback)
author_label.bind("<Enter>", lambda e: author_label.config(cursor="hand2"))
author_label.bind("<Leave>", lambda e: author_label.config(cursor="arrow"))

select_button = Button(root, text="Select Excel Files", command=select_files)
select_button.pack(pady=5)

status_label = Label(root, text="")
status_label.pack()

executor = concurrent.futures.ThreadPoolExecutor()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()
