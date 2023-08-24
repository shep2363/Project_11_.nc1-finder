import os
import shutil
import pandas as pd
from tkinter import filedialog
from tkinter import Tk, Button, Label, Entry, StringVar


def browse_directory():
    filename = filedialog.askdirectory()
    directory_var.set(filename)


def browse_new_directory():
    filename = filedialog.askdirectory()
    new_dir_var.set(filename)


def browse_file():
    filename = filedialog.askopenfilename(filetypes=(("Excel files", "*.xlsx;*.xls"), ("all files", "*.*")))
    file_var.set(filename)


def process_files():
    directory_to_search = directory_var.get()
    excel_file_path = file_var.get()
    new_folder_path = new_dir_var.get()

    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)

    df = pd.read_excel(excel_file_path)

    if "Piece mark" in df.columns:
        column_name = "Piece mark"
    elif "piece mark" in df.columns:
        column_name = "piece mark"
    else:
        print("Neither Columns were found")
        return

    missing_nc1_files = []

    # Test for file read errors
    for file in os.listdir(directory_var.get()):
        try:
            loadFile = open(os.path.join(directory_var.get(), file), 'rb')
            print('.', end='')
            loadFile.read(1)
            loadFile.close()
        except:
            print('Error opening file %s' % file)

    # Process .nc1 files
    for ref_number in df[column_name]:
        for filename in os.listdir(directory_to_search):
            if filename.lower().endswith('.nc1'):
                file_last_5_chars = os.path.splitext(filename)[0][-5:]
                if str(ref_number) == file_last_5_chars:
                    source_path = os.path.join(directory_to_search, filename)
                    dest_path = os.path.join(new_folder_path, filename)
                    shutil.copy(source_path, dest_path)

    print("Process Completed!")



root = Tk()

directory_var = StringVar()
file_var = StringVar()
new_dir_var = StringVar()

directory_entry = Entry(root, textvariable=directory_var, width=50)
directory_entry.pack()
directory_button = Button(root, text="Browse for .nc1 Directory", command=browse_directory)
directory_button.pack()

file_entry = Entry(root, textvariable=file_var, width=50)
file_entry.pack()
file_button = Button(root, text="Browse for Excel File", command=browse_file)
file_button.pack()

new_dir_entry = Entry(root, textvariable=new_dir_var, width=50)
new_dir_entry.pack()
new_dir_button = Button(root, text="Browse for New Directory", command=browse_new_directory)
new_dir_button.pack()

process_button = Button(root, text="Process Files", command=process_files)
process_button.pack()

root.mainloop()
