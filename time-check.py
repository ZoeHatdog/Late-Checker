# -*- coding: utf-8 -*-
import pythoncom
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime
import win32com.client as win32
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, ttk, Label
import customtkinter as ctk
from PIL import Image, ImageTk, ImageDraw
from PIL import Image, ImageTk, ImageDraw

def decimal_to_time(decimal):
    total_seconds = decimal * 24 * 3600
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    return datetime.strptime(f"{hours}:{minutes}", "%H:%M").strftime("%I:%M %p")

def calculate_minutes_late(actual_time_str, expected_time_str):
    try:
        actual_time = datetime.strptime(actual_time_str, "%I:%M %p")
        expected_time = datetime.strptime(expected_time_str, "%I:%M %p")
        difference = (actual_time - expected_time).total_seconds() / 60
        return max(0, difference)  # Ensure the difference is not negative
    except ValueError:
        # If the time string is empty or not in the expected format, return 0 minutes late
        return 0

def modify_columns(xlsx_file_path, sheet_name, time_in_col, time_out_col, work_hours_col, minutes_late_col):
    wb = openpyxl.load_workbook(xlsx_file_path)
    ws = wb[sheet_name]

    headers = [cell.value for cell in ws[1]]
    time_in_index = headers.index(time_in_col) + 1
    time_out_index = headers.index(time_out_col) + 1
    work_hours_index = headers.index(work_hours_col) + 1
    minutes_late_index = headers.index(minutes_late_col) + 1

    for row in ws.iter_rows(min_row=2, max_col=max(time_in_index, time_out_index, work_hours_index, minutes_late_index)):
        cell_time_in = row[time_in_index - 1]
        cell_time_out = row[time_out_index - 1]
        cell_work_hours = row[work_hours_index - 1]
        cell_minutes_late = row[minutes_late_index - 1]

        if isinstance(cell_work_hours.value, (float, int)):
            work_hours_rounded = round(cell_work_hours.value)
            if work_hours_rounded == 9:
                cell_time_in.value = "08:00 AM"
                cell_time_out.value = "05:00 PM"
                cell_minutes_late.value = 0
            elif work_hours_rounded == 12:
                cell_time_in.value = "07:00 PM"
                cell_time_out.value = "07:00 AM"
                cell_minutes_late.value = calculate_minutes_late(cell_time_in.value, "07:00 PM")
            else:
                if isinstance(cell_time_in.value, (float, int)):
                    cell_time_in.value = decimal_to_time(cell_time_in.value)
                if isinstance(cell_time_out.value, (float, int)):
                    cell_time_out.value = decimal_to_time(cell_time_out.value)
                if isinstance(cell_time_in.value, str) and cell_time_in.value.strip():
                    cell_minutes_late.value = calculate_minutes_late(cell_time_in.value, "08:00 AM")

    wb.save(xlsx_file_path)

def xls_to_xlsx(xls_file_path, xlsx_file_path):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False

    # Open the .xls file
    wb = excel.Workbooks.Open(os.path.abspath(xls_file_path))

    # Save it as .xlsx
    wb.SaveAs(os.path.abspath(xlsx_file_path), FileFormat=51)  # 51 is the file format for .xlsx
    wb.Close()
    excel.Quit()

    print(f"Conversion complete. Saved as {xlsx_file_path}")

def process_excel_files(input_file_path, output_file_path):
    xlsx_temp_file = 'temp_converted.xlsx'
    sheet_name = 'Timekeep'
    time_in_col = 'TIME IN'
    time_out_col = 'TIME OUT'
    work_hours_col = 'WORK HOURS'
    minutes_late_col = 'MINUTES LATE'
    
    xls_to_xlsx(input_file_path, xlsx_temp_file)
    modify_columns(xlsx_temp_file, sheet_name, time_in_col, time_out_col, work_hours_col, minutes_late_col)
    os.rename(xlsx_temp_file, output_file_path)

def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    input_entry.delete(0, ctk.END)
    input_entry.insert(0, file_path)

def select_output_file(input_file_path):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            process_excel_files(input_file_path, file_path)
            messagebox.showinfo("Success", f"File processed successfully and saved as {file_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

def start_processing():
    input_file_path = input_entry.get()

    if not input_file_path:
        messagebox.showerror("Error", "Please select an input file")
        return

    select_output_file(input_file_path)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def time_diff_calculator():
    calculator_window = tk.Toplevel(app)  # Create a new window
    calculator_window.title("Time Diff Calculator")
    
    # Set dimensions for the calculator window
    window_width = 400
    window_height = 300
    calculator_window.geometry(f"{window_width}x{window_height}")

    # Get the position of the main window
    main_window_width = app.winfo_width()
    main_window_height = app.winfo_height()
    main_window_x = app.winfo_x()
    main_window_y = app.winfo_y()
    
    # Calculate the center position for the new window
    position_x = main_window_x + (main_window_width - window_width) // 2
    position_y = main_window_y + (main_window_height - window_height) // 2
    
    calculator_window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    def calculate_time_difference():
        try:
            time1 = time1_entry.get()
            time1_hour, time1_minute = map(int, time1.split(":"))
            time1_period = time1_period_var.get()
            if time1_period == "PM" and time1_hour != 12:
                time1_hour += 12
            elif time1_period == "AM" and time1_hour == 12:
                time1_hour = 0

            time2 = time2_entry.get()
            time2_hour, time2_minute = map(int, time2.split(":"))
            time2_period = time2_period_var.get()
            if time2_period == "PM" and time2_hour != 12:
                time2_hour += 12
            elif time2_period == "AM" and time2_hour == 12:
                time2_hour = 0

            time1_total_minutes = time1_hour * 60 + time1_minute
            time2_total_minutes = time2_hour * 60 + time2_minute

            minute_difference = abs(time2_total_minutes - time1_total_minutes)

            result_label.configure(text=f"The minute difference is: {minute_difference} minutes")
        except ValueError:
            messagebox.showerror("Error", "Please enter valid time values")

    # Create a main frame to hold all the widgets
    main_frame = ctk.CTkFrame(calculator_window, fg_color="white")
    main_frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Create a title label
    title_label = ctk.CTkLabel(main_frame, text="Time Calculator", font=ctk.CTkFont(size=24, weight="bold"), text_color="#000000")
    title_label.pack(pady=10)

    # Create a frame for time 1
    time1_frame = ctk.CTkFrame(main_frame, fg_color="white")
    time1_frame.pack(pady=10)

    time1_label = ctk.CTkLabel(time1_frame, text="Time 1:", font=ctk.CTkFont(size=16), text_color="#000000")
    time1_label.pack(side=tk.LEFT, padx=5)

    time1_entry = ctk.CTkEntry(time1_frame, width=60)
    time1_entry.pack(side=tk.LEFT, padx=5)

    time1_period_var = StringVar()
    time1_period_menu = ttk.Combobox(time1_frame, textvariable=time1_period_var, width=15)
    time1_period_menu['values'] = ("AM", "PM")
    time1_period_menu.current(0)
    time1_period_menu.pack(side=tk.LEFT, padx=5)

    # Create a frame for time 2
    time2_frame = ctk.CTkFrame(main_frame, fg_color="white")
    time2_frame.pack(pady=10)

    time2_label = ctk.CTkLabel(time2_frame, text="Time 2:", font=ctk.CTkFont(size=16), text_color="#000000")
    time2_label.pack(side=tk.LEFT, padx=5)

    time2_entry = ctk.CTkEntry(time2_frame, width=60)
    time2_entry.pack(side=tk.LEFT, padx=5)

    time2_period_var = StringVar()
    time2_period_menu = ttk.Combobox(time2_frame, textvariable=time2_period_var, width=15)
    time2_period_menu['values'] = ("AM", "PM")
    time2_period_menu.current(0)
    time2_period_menu.pack(side=tk.LEFT, padx=5)

    # Create a calculate button
    calculate_button = ctk.CTkButton(main_frame, text="Calculate", font=ctk.CTkFont(size=16, weight="bold"), command=calculate_time_difference)
    calculate_button.pack(pady=10)

    # Create a result label
    result_label = ctk.CTkLabel(main_frame, text="", font=ctk.CTkFont(size=16), text_color="#000000")
    result_label.pack(pady=10)

# Main application window


app = ctk.CTk()
app.geometry("1000x600")
app.title("Time Processor App")

image_path = resource_path("imgs/time.png")
img = Image.open(image_path)

background = Image.new("RGBA", img.size, (255, 255, 255, 0))  # Transparent background
background.paste(img, (0, 0), img)
background = background.convert("RGB")

img = background.resize((200, 200), Image.LANCZOS)
img = ImageTk.PhotoImage(img)
img_label = Label(app, image=img, bg='white', borderwidth=0, highlightthickness=0)
img_label.pack(pady=10)

frame = ctk.CTkFrame(app, width=900, height=500, corner_radius=10, fg_color="white")
frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

label_title = ctk.CTkLabel(frame, text="HR Data Processor", font=ctk.CTkFont(size=24, weight="bold"), text_color="#000000")
label_title.pack(pady=20)

input_frame = ctk.CTkFrame(frame, width=800, height=100, fg_color="white")
input_frame.pack(pady=10)

input_label = ctk.CTkLabel(input_frame, text="Input Excel File:", font=ctk.CTkFont(size=16), text_color="#000000")
input_label.pack(side=tk.LEFT, padx=10)

input_entry = ctk.CTkEntry(input_frame, width=400)
input_entry.pack(side=tk.LEFT, padx=10)

input_button = ctk.CTkButton(input_frame, text="Browse", command=select_input_file, fg_color="#3498db")
input_button.pack(side=tk.LEFT, padx=10)

process_button = ctk.CTkButton(frame, text="Process", command=start_processing, fg_color="#2ecc71", font=ctk.CTkFont(size=16, weight="bold"))
process_button.pack(pady=20)

# Time Diff Calculator button
time_diff_button = ctk.CTkButton(frame, text="Time Diff Calculator", command=time_diff_calculator, fg_color="#e74c3c", font=ctk.CTkFont(size=16, weight="bold"))
time_diff_button.pack(pady=10)

app.mainloop()
