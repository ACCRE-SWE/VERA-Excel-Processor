import pandas as pd
import re
import sys


SUMMARY = "SUMMARY"
SUMMARY_PERSONNEL = "SUMMARY_PERSONNEL"
SUMMARY_GENERAL = "SUMMARY_GENERAL"
SUMMARY_SUBAWARDS = "SUMMARY_SUBAWARDS"

SPONSOR = "SPONSOR"
SPONSOR_PERSONNEL = "SPONSOR_PERSONNEL"
SPONSOR_GENERAL = "SPONSOR_SUMMARY_GENERAL"
SPONSOR_SUBAWARDS = "SPONSOR_SUMMARY_SUBAWARDS"


COST_SHARE_SUMMARY ="COST_SHARE_SUMMARY"
COST_SHARE_SUMMARY_PERSONNEL = "COST_SHARE_SUMMARY_PERSONNEL"
COST_SHARE_SUMMARY_GENERAL = "COST_SHARE_SUMMARY_GENERAL"
COST_SHARE_SUMMARY_SUBAWARDS = "COST_SHARE_SUMMARY_SUBAWARD"

FUNDS_REQUESTED = "FUNDS_REQUESTED"
FUNDS_REQUESTED_PERSONNEL = "FUNDS_REQUESTED_PERSONNEL"
FUNDS_REQUESTED_GENERAL = "FUNDS_REQUESTED_GENERAL"
FUNDS_REQUESTED_SUBAWARDS = "FUNDS_REQUESTED_SUBAWARD"

COST_SHARE ="COST_SHARE"
COST_SHARE_PERSONNEL = "COST_SHARE_PERSONNEL"
COST_SHARE_GENERAL = "COST_SHARE_GENERAL"
COST_SHARE_SUBAWARDS = "COST_SHARE_SUBAWARD"


#CONSTS
PERSONNEL = "Personnel"
GENERAL = "General"
SUBAWARD = "Subaward"

LINE_BREAK_SECTIONS = [
    SUMMARY_GENERAL,
    SUMMARY_SUBAWARDS,
    SPONSOR_GENERAL,
    SPONSOR_SUBAWARDS,
    COST_SHARE_SUMMARY_GENERAL,
    COST_SHARE_SUMMARY_SUBAWARDS,
    FUNDS_REQUESTED_GENERAL,
    FUNDS_REQUESTED_SUBAWARDS,
    COST_SHARE_GENERAL,
    COST_SHARE_SUBAWARDS,
]


def get_current_section(current_section, row):
    val = row.values[0]
    if pd.isnull(val):
        return current_section
    val = str(val)
    if not current_section and "Project Overall Summary:" in val:
        current_section = SUMMARY
    elif current_section == SUMMARY and  PERSONNEL in val:
        current_section = SUMMARY_PERSONNEL
    elif current_section == SUMMARY_PERSONNEL and  GENERAL in val:
        current_section = SUMMARY_GENERAL
    elif current_section == SUMMARY_GENERAL and SUBAWARD in val:
        current_section = SUMMARY_SUBAWARDS
    
    elif current_section == SUMMARY_SUBAWARDS and "Sponsor Summary:" in val:
        current_section = SPONSOR
    elif current_section == SPONSOR and  PERSONNEL in val:
        current_section = SPONSOR_PERSONNEL
    elif current_section == SPONSOR_PERSONNEL and GENERAL in val:
        current_section = SPONSOR_GENERAL
    elif current_section == SPONSOR_GENERAL and SUBAWARD in val:
        current_section = SPONSOR_SUBAWARDS

    elif current_section == SPONSOR_SUBAWARDS and "Cost Share Summary:" in val:
        current_section = COST_SHARE_SUMMARY
    elif current_section == COST_SHARE_SUMMARY and  PERSONNEL in val:
        current_section = COST_SHARE_SUMMARY_PERSONNEL
    elif current_section == COST_SHARE_SUMMARY_PERSONNEL and  GENERAL in val:
        current_section = COST_SHARE_SUMMARY_GENERAL
    elif current_section == COST_SHARE_SUMMARY_GENERAL and SUBAWARD in val:
        current_section = COST_SHARE_SUMMARY_SUBAWARDS
    
    elif "Funds Requested" in val:
        current_section = FUNDS_REQUESTED
    elif current_section == FUNDS_REQUESTED and  PERSONNEL in val:
        current_section = FUNDS_REQUESTED_PERSONNEL
    elif current_section == FUNDS_REQUESTED_PERSONNEL and  GENERAL in val:
        current_section = FUNDS_REQUESTED_GENERAL
    elif current_section == FUNDS_REQUESTED_GENERAL and SUBAWARD in val:
        current_section = FUNDS_REQUESTED_SUBAWARDS

      
    elif "Cost Share" in val:
        current_section = COST_SHARE
    elif current_section == COST_SHARE and  PERSONNEL in val:
        current_section = COST_SHARE_PERSONNEL
    elif current_section == COST_SHARE_PERSONNEL and  GENERAL in val:
        current_section = COST_SHARE_GENERAL
    elif current_section == COST_SHARE_GENERAL and SUBAWARD in val:
        current_section = COST_SHARE_SUBAWARDS
        
    return current_section
    
def parse_personnel(data):
    """
        Params 
        data : Pass in row array
    """
    try:
        name_prefix = ",".join(seg.strip() for seg in data[0].split('\n'))
        rows = ["Effort","FBRate","Base","Salary","Benefits","Total"]
        good_data = []
        for idx in range(1,len(data)):
            cell_val = data[idx]
            if pd.isnull(cell_val):
                good_data.append({})
                continue

            reg = ":\s(.*?)(\\n|$)"
            obj = {}
            for row in rows:
                matc = re.findall(row + reg, cell_val)
                if matc:
                    obj[row] = matc[0][0]
            good_data.append(obj)

        res = []
        for row in rows:
            r = [name_prefix + "," + row]
            for col in good_data:
                if row in col:
                    r.append(col[row])
                else:
                    r.append("")
            res.append(r)
        return res
    except Exception:
        pass

def process_line_breaks(row):
    "must return a list of dataframes"
    retval = []
    for val in row.values:
        proc_val = None
        if pd.isnull(val):
            proc_val = ""

        proc_val = str(val)
        for i, sp in enumerate(proc_val.split("\n")):
            try:
                retval[i].append(sp)
            except IndexError:
                retval.append([])
                retval[i].append(sp)

    return retval


def formatter(file_path):
    df = pd.read_excel(file_path,header=None)
    df.fillna("",inplace=True)
    change_copy = df.copy(deep=True)
    skip_next = False
    current_section, old_section = None, None
    itr = df.iterrows()
    index_offset = 0

    while True:
        try:
            if skip_next:
                skip_next = False
            else:
                index, row = next(itr)

            # print(row.values)
            current_section = get_current_section(current_section, row)
            # print(current_section, row.values)

            if current_section in LINE_BREAK_SECTIONS and current_section != old_section:
                index, row = next(itr)
                new_rows = process_line_breaks(row)
                append_df = pd.DataFrame(new_rows, columns=df.columns)
                change_copy = pd.concat([change_copy.iloc[:index + index_offset], append_df, change_copy.iloc[index + index_offset + 1:]], ignore_index=True)
                change_copy.reset_index()
                index_offset += len(new_rows) - 1
            
            if current_section == FUNDS_REQUESTED_PERSONNEL or current_section == COST_SHARE_PERSONNEL:
                index, row = next(itr) # salaries
                index, row = next(itr) # benifits
                index, row = next(itr)
                while "Person:" in row.values[0]:
                    new_rows = parse_personnel(row.values)
                    append_df = pd.DataFrame(new_rows, columns=df.columns)
                    change_copy = pd.concat([change_copy.iloc[:index + index_offset], append_df, change_copy.iloc[index + index_offset + 1:]], ignore_index=True)
                    change_copy.reset_index()
                    index_offset += len(new_rows) - 1
                    index, row = next(itr)
                    change_copy.to_csv("output.csv")
                skip_next = True

            old_section = current_section
        except StopIteration:
            break

    return change_copy

try:
    import Tkinter as tk
    from tkFileDialog import askopenfilename
    from tkMessageBox import showinfo, showerror
except ModuleNotFoundError:   # Python 3
    import tkinter as tk
    from tkinter.filedialog import askopenfilename
    from tkinter.messagebox import showinfo, showerror

import os

root = tk.Tk()
root.title('VERA Excel Transformer')
root.geometry("500x150")

def open_file():
    try:
        filename = askopenfilename(filetypes=[("Excel files","*.xlsx")])
        df = formatter(filename)
        name = filename.split(".xlsx")[0]
        df.to_excel(f"{name}_output.xlsx",header=None,index=None)
        showinfo(title="Success", message="File exported successfully. Check alongside the input file.")
    except Exception as error:
        showerror(title="Failed", message=f"Something went wrong. Error: {str(error)}")
        raise error


my_text1 = tk.Label(root, text="Click the button below to select the excel file to process.")
my_text1.pack(pady=5)
my_text2 = tk.Label(root, text="It will create the output file alongside the input with _output appended to filename")
my_text2.pack(pady=5)

my_button = tk.Button(root, text="Open File", command=open_file)
my_button.pack(pady=20)
root.mainloop()
