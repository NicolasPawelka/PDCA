import win32com.client as win32
import tkinter as tk
from tkinter import filedialog,messagebox
import sys
##########################
DEMO = True
EXCEL = None
PROJECT = None
EXCEL_FILE_PATH = None
ACTIVE_PROJECT = None
TASKS = None
WORKBOOK = None
WORKSHEET = None
SHEET_NAME = None
NAME = []
KEY = []
START = []
END = []
PROJECT_ID = []
#########################
def get_positiion():
    global EXCEL
    active_sheet = EXCEL.ActiveSheet
    
    for i, sheet in enumerate(EXCEL.Worksheets, 1):
        if sheet.Name == active_sheet.Name:
            return i - 1
    return None

def choose_project_file():
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(filetypes=[("Project-Dateien", "*.mpp")])
    return file_path

def init(excel_file_path):
    global EXCEL
    global EXCEL_FILE_PATH
    global PROJECT
    global ACTIVE_PROJECT
    global TASKS
    global WORKSHEET
    
    project_file_path = choose_project_file()
    
    PROJECT = win32.Dispatch("MSProject.Application")
    EXCEL = win32.Dispatch("Excel.Application")
    EXCEL_FILE_PATH = excel_file_path
    WORKSHEET = EXCEL.ActiveWorkbook.Sheets("1. Software Entwicklungssupport")
    PROJECT.FileOpen(project_file_path)
    
    ACTIVE_PROJECT = PROJECT.ActiveProject
    TASKS = ACTIVE_PROJECT.Tasks
    main()

def sorting_key(item):
    parts = item.split('.')
    return [int(part) if part.isdigit() else part for part in parts]
    
def main():
    global EXCEL
    global EXCEL_FILE_PATH
    global PROJECT
    global ACTIVE_PROJECT
    global TASKS
    global WORKSHEET
    global KEY
    
    if TASKS is None:
        messagebox.showerror("Error", "Keine Project Daten erhalten")
        sys.exit(1)
        
    for task in TASKS:
        KEY.append(f"{get_positiion()}.{task.OutlineNumber}")
        NAME.append(task.Name)
        START.append(task.Start.strftime('%d.%m.%Y'))
        END.append(task.Finish.strftime('%d.%m.%Y'))
        PROJECT_ID.append(task.GUID)
    if KEY is None or NAME is None or START is None or END is None:
        messagebox.showerror("Error", "DATA ist leer")
        sys.exit(1)
    EXCEL.Application.Run("Transfer",NAME,KEY,START,END,PROJECT_ID)
    messagebox.showinfo("Completed", "Import der Daten erfolgreich!")
     
                

if __name__ == "__main__":
    if DEMO is False:
        if len(sys.argv) != 2:
            messagebox.showerror("Error","Fehler bei der Ermittlung des Commands")
            sys.exit()
        else:
            excel_path = sys.argv[1]
            init(excel_path)
        
    init(r"C:\Users\npawelka\Desktop\PDCA\PDCA_ETO.EEVACTUATOR.Entw.016.xlsm")