import os
import win32com.client
import tkinter as tk
from tkinter import messagebox

# Path to the shortcut file
shortcut_path = r"C:\Users\Packing\Desktop\Nov 24 Shun Shing Cement Industries Ltd - Shortcut.lnk"

# Function to resolve the shortcut path to the actual Excel file path
def resolve_shortcut(shortcut_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    return shortcut.TargetPath

# Path to the target file (resolved from the shortcut)
excel_file_path = resolve_shortcut(shortcut_path)

# Function to check if the Excel file is open
def check_if_file_is_running():
    excel = win32com.client.Dispatch("Excel.Application")
    for wb in excel.Workbooks:
        if wb.FullName.lower() == excel_file_path.lower():
            return True
    return False

# Function to close the specific Excel file
def close_file():
    excel = win32com.client.Dispatch("Excel.Application")
    for wb in excel.Workbooks:
        if wb.FullName.lower() == excel_file_path.lower():
            wb.Close(SaveChanges=False)  # Close the workbook without saving changes
            messagebox.showinfo("Success", f"Excel file '{excel_file_path}' closed successfully!")
            return
    messagebox.showinfo("Not Open", f"The file '{excel_file_path}' is not open in Excel.")

# Creating the tkinter window
root = tk.Tk()
root.withdraw()  # Hide the main window

# Check if the file is running
if check_if_file_is_running():
    if messagebox.askyesno("File Running", f"The Excel file '{excel_file_path}' is currently running. Do you want to close it?"):
        close_file()
else:
    messagebox.showinfo("Not Running", f"The Excel file '{excel_file_path}' is not running.")
