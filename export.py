import subprocess
import sys
import win32com.client as win32
from win32com.client import constants as pjconstants
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from openpyxl import load_workbook
import re
import json
import os
import inspect
############################
DEMO = False
TASK_NAME = "PDCAPosName"
START_DATE = None
END_DATE = "EndDate"
DATA_FRAME = None
EXCEL_FILE_PATH = None
PROJECT = None
ACTIVE_PROJECT = None
SELECTED_SHEET = None
TASKS = None
PROJECT_FILE_PATH = None
WAS_SUMMARY = False
RESOURCES = []
ID = []
DATABASE_NAME = "PDCA_InternDatenbank"
PDCA_POSITION_INDEX = None
EXCEL = None
WORKBOOK = None
############################
def is_summary(current_depth, saved_depth, next_depth):
    if current_depth > saved_depth and current_depth > next_depth:
        return True
    elif current_depth < next_depth:
        return True
    elif current_depth > saved_depth and current_depth == next_depth:
        return True
    return False

def find_existing_task_by_name(name, TASKS):
    for task in TASKS:
        if task.Name == name:
            return task
    return None


def find_Sheet_Name(df):
    num_columns = df.iloc[PDCA_POSITION_INDEX].shape[0]
    for i in range(num_columns):
        if df.iloc[PDCA_POSITION_INDEX,i] == "PDCAPosSheet":
            return i
    return -1

def find_TASK_NAME(df,TASK_NAME):
    num_columns = df.iloc[PDCA_POSITION_INDEX].shape[0]
    for i in range(num_columns):
        if df.iloc[PDCA_POSITION_INDEX,i] == TASK_NAME:
            return i
    return -1

def find_ID(df):
    num_columns = df.iloc[PDCA_POSITION_INDEX].shape[0]
    for i in range(num_columns):
        if df.iloc[PDCA_POSITION_INDEX,i] == "PDCAPosNr":
            return i
    return -1

def find_START(df):
    num_columns = df.iloc[PDCA_POSITION_INDEX].shape[0]
    for i in range(num_columns):
        if df.iloc[PDCA_POSITION_INDEX,i] == "StartDate":
            return i
    return -1

def find_BUDGET(df):
    num_columns = df.iloc[PDCA_POSITION_INDEX].shape[0]
    for i in range(num_columns):
        if df.iloc[PDCA_POSITION_INDEX,i] == "BudgetHoursPlanned":
            return i
    return -1


def find_RESOURCE(df):
    length = len(df.iloc[:,1])
    for i in range(length):
        if df.iloc[i,1] == "Mitarbeiter im Projekt":
            return i
    return -1

def find_PDCA_Positionen(df):
    length = len(df.iloc[:,1])
    for i in range(length):
        if df.iloc[i,1] == "PDCAPosID":
            return i
    return -1

def choose_excel_file():
    root = tk.Tk()
    root.withdraw() # Vertsecke das Hauptfenster

    file_path = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx *.xls *.xlsm")])

    return file_path


def choose_excel_sheet(file_path):
    if file_path:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names

        #Popup Dialog Fenster öffnen um das sheet auszuwählen
        options = ["Load all sheets"]
        options.extend([f"Load sheet {i + 1}: {sheet_name}" for i, sheet_name in enumerate(sheet_names)])

        choice = simpledialog.askinteger("Auswahl der Arbeitsmappe", "Waählen Sie eine Arbeitsmappe die geladen werden soll: ",
                                         initialvalue=0, minvalue=0, maxvalue=len(options)-1)

        choice += 1
        if choice == 0:
           return None
        elif 1 <= choice < len(options):
            return sheet_names[choice-1]
    return None


def calculate_depth(text):
    numbers = re.findall(r'\.', text)
    return len(numbers)

def extract_budget(text):
    pattern = r'\*(\d+)'
    if isinstance(text,int):
        return text
    elif text is None:
        return 0

    match = re.search(pattern,text)

    if match:
        extrcted_number = int(match.group(1))
        return extrcted_number
    return -1

def add_Task(TASKS,name,depth,date,budget,vorgänger):
    global EXCEL
    task = TASKS.Add()
    task.Manual = False
    task.Name = name
    #task.Start = date
    task.OutlineLevel = depth
    task.Cost = budget
    EXCEL.Application.Run("Sync",task.GUID,task.Name)
    return 1

def add_Summary(TASKS, name,depth,date,budget):
    global EXCEL
    task = TASKS.Add()
    task.Manual = False
    task.Name = name
   # task.Start = date
    task.OutlineLevel = depth
    EXCEL.Application.Run("Sync",task.GUID,task.Name)
    return 1

def add_resource():
    RESOURCE_index = find_RESOURCE(DATA_FRAME) + 1

    for _, row in DATA_FRAME.iloc[RESOURCE_index:].iterrows():
        if row.iloc[1] is not None:
            if row.iloc[1] not in [resource.Name for resource in ACTIVE_PROJECT.Resources]:
               resource = ACTIVE_PROJECT.Resources.Add(row.iloc[1])
        else:
            break

def extract_file_name(text):
    match = re.search(r'[^\\/]+$', text)

    if match:
        filename = match.group(0)
        return filename
    else:
        messagebox.showerror("Error", "Name der verbundenen Datei konnte nicht ermittelt werden")

def find_file_path(name):
    command = f'cmd /c "cd ../.. & cd Desktop & dir /s /b "{name}"'

    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,shell=True,text=True)
    stdout = process.communicate()
    stdout = r'C:\Users\npawelka\Desktop\PDCA\PDCA_ETO.EEVACTUATOR.Entw.016.xlsm'
    
    return stdout

def convert_to_raw_string(input):
   return r'{}'.format(input[:-1])

def init(project_file_path):
    global DATA_FRAME
    global EXCEL_FILE_PATH
    global PROJECT
    global SELECTED_SHEET
    global ACTIVE_PROJECT
    global TASKS
    global PROJECT_FILE_PATH
    global ID
    global DATABASE_NAME
    global PDCA_POSITION_INDEX
    global RESOURCES
    global START_DATE
    global EXCEL
    global WORKBOOK
    
    calling_function = inspect.currentframe().f_back.f_code.co_name

    if calling_function != "update":
        EXCEL_FILE_PATH = choose_excel_file()
        SELECTED_SHEET = choose_excel_sheet(EXCEL_FILE_PATH)
        EXCEL = win32.Dispatch("Excel.Application")
        PROJECT = win32.Dispatch("MSProject.Application")
        PROJECT_FILE_PATH = project_file_path
        PROJECT.FileOpen(project_file_path)
        script_directory = os.path.dirname(os.path.abspath(__file__))
        json_filename = os.path.join(script_directory,"connected_values.json")


        if not os.path.exists(json_filename):
            data = {}
        else:
            with open(json_filename,"r") as json_file:
                data = json.load(json_file)


        new_key = extract_file_name(PROJECT_FILE_PATH)
        new_value = [extract_file_name(EXCEL_FILE_PATH), SELECTED_SHEET[0]]
        data[new_key] = new_value

        with open(json_filename, "w") as json_file:
            json.dump(data, json_file, indent=4)
    else:
         with open("connected_values.json", "r") as json_file:
            data = json.load(json_file)
         key = extract_file_name(project_file_path)
         EXCEL_FILE_PATH = find_file_path(data[key][0])
         EXCEL_FILE_PATH = convert_to_raw_string(EXCEL_FILE_PATH)
         SELECTED_SHEET = data[key][1]
         PROJECT_FILE_PATH = project_file_path
         PROJECT = win32.Dispatch("MSProject.Application")
         PROJECT.FileOpen(PROJECT_FILE_PATH)

         
    workbook = load_workbook(EXCEL_FILE_PATH, data_only=True)
    WORKBOOK = workbook
    worksheet = workbook[DATABASE_NAME]

    DATA_FRAME = pd.DataFrame(worksheet.values)
    DATA_FRAME = DATA_FRAME.reset_index()
    PDCA_POSITION_INDEX = find_PDCA_Positionen(DATA_FRAME)

    ACTIVE_PROJECT = PROJECT.ActiveProject
    TASKS = ACTIVE_PROJECT.Tasks

    ID_index = find_ID(DATA_FRAME)
    last_value = None
    for _, row in DATA_FRAME.iloc[PDCA_POSITION_INDEX:].iterrows():
        if row.iloc[ID_index] is not None:
            last_value = row.iloc[ID_index]
            ID.append(last_value)

    if last_value is not None:
        ID.append(last_value)
        
    add_resource()
    
    START_DATE = workbook['Cockpit']['F16'].value
    
    if calling_function == "update":
        update("SUCCESS")
    else:
        main()

def update(mpp_file_path):
    global DATA_FRAME
    global EXCEL_FILE_PATH
    global PROJECT
    global ACTIVE_PROJECT
    global TASKS


    if  DATA_FRAME is None or EXCEL_FILE_PATH is None or PROJECT is None or ACTIVE_PROJECT is None:
        init(mpp_file_path)
    else:
        Task_Name_index = find_TASK_NAME(DATA_FRAME, TASK_NAME)
        ID_index = find_ID(DATA_FRAME)
        START_index = find_START(DATA_FRAME)
        BUDGET_index = find_BUDGET(DATA_FRAME)
    
        if Task_Name_index == -1 or ID_index == -1 or START_index == -1 or BUDGET_index == -1:
            messagebox.showerror("Error", "Could not find required columns in the Excel data.")
            sys.exit(1)
    
        current_index = 1
        for _, row in DATA_FRAME.iterrows():
            current_name = row.iloc[Task_Name_index]
            current_id = row.iloc[ID_index]
            current_budget = row.iloc[BUDGET_index]
            if current_name is None or current_name == TASK_NAME or current_name == SELECTED_SHEET:
                continue
            else:
                current_depth = calculate_depth(current_id)
                date = row.iloc[START_index]
                if date is None:
                    date = START_DATE
                task = find_existing_task_by_name(current_name, TASKS)
                if task:
                    task.Start = date
                    task.OutlineLevel = current_depth
                    #task.Cost = extract_budget(current_budget)
                else:
                    messagebox.showwarning("Warning", f"Task '{current_name}' not found in the existing project. Skipping.")
                current_index += 1
    
        messagebox.showinfo("Update Complete", "Existing tasks have been updated.")


def main():
    global ID
    global START_DATE
    global WORKBOOK
    current_index = 1
    saved_depth = 0
    task_number = 0
    first_summary = True
    Task_Name_index = find_TASK_NAME(DATA_FRAME, TASK_NAME)
    ID_index = find_ID(DATA_FRAME)
    START_index = find_START(DATA_FRAME)
    BUDGET_index = find_BUDGET(DATA_FRAME)
    SHEET_index = find_Sheet_Name(DATA_FRAME)
    if Task_Name_index == -1 or ID_index == -1 or START_index == -1 or BUDGET_index == -1 or SHEET_index == -1:
        messagebox.showerror("Error", "Could not find Column")
        sys.exit()


    for index,row in DATA_FRAME.iloc[PDCA_POSITION_INDEX:].iterrows():
        global TASKS
        global WAS_SUMMARY

        current_name = row.iloc[Task_Name_index]
        current_id = row.iloc[ID_index]
        current_budget = row.iloc[BUDGET_index]
        current_sheet = row.iloc[SHEET_index]
        if current_name is None or current_name == TASK_NAME or current_name == SELECTED_SHEET or current_sheet != SELECTED_SHEET:
            continue
        else:
            current_depth = calculate_depth(current_id)
            next_depth = calculate_depth(ID[current_index + 1])
            current_index += 1
            if current_depth == 0:
                continue
            if first_summary:
                task_number += 1
                first_summary = False
                add_Summary(TASKS,current_name,current_depth,START_DATE,extract_budget(current_budget))
                WAS_SUMMARY = True
                saved_depth = current_depth
            else:
                date = WORKBOOK[DATABASE_NAME]['K'][index].value
                if date == None:
                    date = START_DATE
                if is_summary(current_depth,saved_depth,next_depth) :
                    add_Summary(TASKS,current_name,current_depth,date,extract_budget(current_budget))
                    task_number += 1
                    WAS_SUMMARY = True
                    saved_depth = current_depth
                else:
                        add_Task(TASKS,current_name,current_depth,date,extract_budget(current_budget),task_number)
                        task_number += 1
                        saved_depth = current_depth


    #project.FileSave()
    TASKS = None
    messagebox.showinfo("Completed!","Import der Daten erfolgreich")

if __name__ == "__main__":
    if DEMO is False:
        if len(sys.argv) != 3:
            messagebox.showerror("Error", "Fehler bei der Ermittlung des Commands")
            sys.exit()
        else:
            if sys.argv[2] == "update":
                update(sys.argv[1])
            else:
                mpp_file_path = sys.argv[1]
                init(mpp_file_path)
    else:
        mpp_file_path = r"C:\Users\npawelka\Desktop\PDCA_WorkingCopy.mpp"
        init(mpp_file_path)
