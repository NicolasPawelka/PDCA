import shutil
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
import sys
from tkinter import messagebox

def copy_mpp_file():
    root = tk.Tk()
    root.withdraw()

    source_file_path = sys.argv[1]

    if not source_file_path:
        print("No source file selected. Exiting.")
        return

    destination_directory = filedialog.askdirectory(
        title="Select Destination Directory"
    )

    if not destination_directory:
        return
   
    destination_path = f"{destination_directory}/PDCA_WorkingCopy.mpp"
    try:
        shutil.copy(source_file_path, destination_path)
        print(f"File copied successfully from {source_file_path} to {destination_path}")
    except FileNotFoundError:
        print("File not found. Please provide valid file paths.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    copy_mpp_file()