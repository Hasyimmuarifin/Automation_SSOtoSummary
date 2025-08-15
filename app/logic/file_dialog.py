import tkinter as tk
from tkinter import filedialog

def select_excel_file():
    """
    Open a file dialog to allow the user to select an Excel file.
    Returns the full path of the selected file.
    Raises an exception if no file is selected.
    """
    # Initialize the hidden root window for the dialog
    root = tk.Tk()
    root.withdraw()

    # Open the file dialog to select an Excel file
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    # Handle case when user cancels the file selection
    if not file_path:
        raise Exception("❌ No file selected. Program terminated.")

    return file_path