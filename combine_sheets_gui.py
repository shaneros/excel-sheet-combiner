import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import datetime

app = tk.Tk()
app.title('Sheet Combiner')
app.geometry("640x480")

instructions_text = """
Click the 'Choose Excel Workbook' button and select the file that you'd like to combine.
When you've confirmed your choice, click the 'Combine Sheets' button.
The output file will be in the same directory as this executable.
"""
allowed_filetypes = [('Excel files', '*.xls *.xlsx'), ('all files', '*.*')]
app.file_label = tk.StringVar()
app.new_file_label = tk.StringVar()

app.file_label.set("No file selected.")
app.new_file_label.set('')

def open_file():
    app.filename = filedialog.askopenfilename(parent = app,
                                                initialdir = os.getcwd(),
                                                title = "Please select a file:",
                                                filetypes = allowed_filetypes)
    app.file_label.set(path_leaf(app.filename))

def client_exit():
    exit()

def path_leaf(path):
    head, tail = os.path.split(path)
    return tail or os.path.basename(head)

def create_output_filename(original_filename):
    now=datetime.datetime.now()
    strs = []
    title, ext = os.path.splitext(path_leaf(original_filename))
    strs.append(title)
    strs.append(' - combined output - ')
    strs.append(now.strftime("%Y-%m-%d %H%Mh"))
    strs.append(ext)
    output_filename = ''.join(strs)
    return output_filename

def combine():
    df = pd.concat(pd.read_excel(app.filename, sheet_name=None), ignore_index=True)
    output = create_output_filename(app.filename)
    df.to_excel(output, index=False)
    app.new_file_label.set(output + ' has been created.')

tk.Label(app, text = instructions_text).pack()

open_btn = tk.Button(app, text = "Choose Excel Workbook", command = open_file).pack()
tk.Label(app, textvariable=app.file_label).pack()
convert_btn = tk.Button(app, text = "Combine Sheets", command = combine).pack()
tk.Label(app, textvariable=app.new_file_label).pack()

#quitButton = tk.Button(app, text="Quit", command=client_exit)
#quitButton.pack()
app.mainloop()