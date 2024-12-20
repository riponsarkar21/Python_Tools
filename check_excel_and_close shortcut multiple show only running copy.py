import os
import win32com.client
import tkinter as tk
from tkinter import messagebox

# Define the source folder and the list of shortcut files
source_folder = r"ProductionFile\2024"
source_files = [
    "Jan 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Feb 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Mar 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Apr 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "May 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Jun 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Jul 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Aug 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Sep 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Oct 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Nov 24 Shun Shing Cement Industries Ltd - Shortcut.lnk",
    "Dec 24 Shun Shing Cement Industries Ltd - Shortcut.lnk"
]

# Function to resolve the shortcut path to the actual Excel file path
def resolve_shortcut(shortcut_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    return shortcut.TargetPath

# Function to check if the Excel file is open
def check_if_file_is_running(excel_file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    for wb in excel.Workbooks:
        if wb.FullName.lower() == excel_file_path.lower():
            return True
    return False

# Function to close the specific Excel file
def close_file(excel_file_path):
    excel = win32com.client.Dispatch("Excel.Application")
    for wb in excel.Workbooks:
        if wb.FullName.lower() == excel_file_path.lower():
            wb.Close(SaveChanges=False)  # Close the workbook without saving changes
            return True
    return False

# Function to process each shortcut
def process_shortcut_files():
    # Creating the tkinter window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Flag to check if any file is running
    any_file_running = False

    for shortcut_file in source_files:
        shortcut_path = os.path.join(source_folder, shortcut_file)

        if os.path.exists(shortcut_path):
            excel_file_path = resolve_shortcut(shortcut_path)

            # Check if the file is running
            if check_if_file_is_running(excel_file_path):
                any_file_running = True
                if messagebox.askyesno("File Running", f"The Excel file '{excel_file_path}' is currently running. Do you want to close it?"):
                    if close_file(excel_file_path):
                        messagebox.showinfo("Success", f"Excel file '{excel_file_path}' closed successfully!")
                    else:
                        messagebox.showinfo("Not Open", f"The file '{excel_file_path}' is not open in Excel.")
    
    # If no file was running
    if not any_file_running:
        messagebox.showinfo("No Files Running", "No Excel files from the list are currently running.")

# Call the function to process the shortcuts
process_shortcut_files()
