import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import CorcManagerExport_multiFilter
import sys

# Functions
def read_params_from_file(file_path):
    params = {}
    with open(file_path, 'r') as file:
        for line in file:
            key, value = line.strip().split('=')
            params[key] = value
    return params

def start_export_button():
    params_file_path = 'parameters.txt'
    params = read_params_from_file(params_file_path)
    
    version = params.get('version_filter', '')
    verdict = params.get('verdict_filter', '')
    target_path = params.get('target_path', '')
    
    CorcManagerExport_multiFilter.main(version, verdict, target_path)
    start_export.configure(text='Export started')

def target_folder():
    target_folder = filedialog.askdirectory()
    target_folder_dir.set(target_folder)
    Select_Target_Folder_Button.configure(text='Selected')
    Target_Folder_label.configure(text=target_folder)

def redirector(inputStr):
    textbox.insert(INSERT, inputStr)

# MAIN
# Create Window
window = Tk()
window.title('Report Manager Export')

# Allow redirect std print to text box
sys.stdout.write = redirector  # type: ignore

# Global variable for target folder path
target_folder_dir = tk.StringVar()

# Create Container Frame for better layout control
frame_controls = tk.LabelFrame(window, text='Controls')

# Target Directory Button
Select_Target_Folder_Button = Button(frame_controls, text='Select Folder to save export', command=target_folder)
Target_Folder_label = Label(frame_controls, text='No Target set:')

# Start Export Button
start_export = Button(frame_controls, text='Start export', command=start_export_button)

# Exit Button
quit_button = Button(window, text='Quit', command=window.quit)

# Console Output Textbox
textbox = Text(window, width=80, height=20)

# Grid placement of all elements
frame_controls.grid(sticky='w')
Select_Target_Folder_Button.grid(row=0, column=0, padx=20, pady=10)
Target_Folder_label.grid(row=0, column=1, padx=20, pady=10)
start_export.grid(row=1, column=0, padx=20, pady=10)
quit_button.grid(row=2, column=0, padx=20, pady=10)
textbox.grid(row=3, column=0, columnspan=2, padx=20, pady=10)

window.mainloop()
