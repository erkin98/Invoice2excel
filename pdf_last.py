#!/usr/bin/env python
# coding: utf-8


import pdfplumber
from tkinter import filedialog
import time
import tkinter
import pandas as pd
import PySimpleGUI as sg
import os.path
import re

# First the window layout in 2 columns

file_list_column = [
    [
        sg.Text("Pdf folder"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse(),
    ],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
        )
    ],
]

# For now will only show the name of the file that was chosen
image_viewer_column = [
    [sg.Text("Choose an pdf from list on left:")],
    [sg.Text(size=(40, 1), key="-TOUT-")],
    [sg.Image(key="-IMAGE-")],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(image_viewer_column),
        
    ],
    [sg.Submit(key='-SUBMIT-')]
     
]

window = sg.Window("Pdf browser", layout)

# Run the Event Loop
while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    # Folder name was filled in, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []

        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            and f.lower().endswith((".pdf"))
        ]
        window["-FILE LIST-"].update(fnames)
    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
            window["-TOUT-"].update(filename)
            window["-IMAGE-"].update(filename=filename)
            
            

        except:
            pass
            

    if event == '-SUBMIT-':
        try:
            filename
        except NameError:
            sg.popup_error("Pdf seçiniz",keep_on_top=True)
            continue        
        else:
            sg.Popup('Pdf dosyası başarıyla seçildi', keep_on_top=True)
            time.sleep(1)
            break
        
window.close()

pdf = pdfplumber.open(filename)
fat = pd.DataFrame()
tables = []
leng = 0
for i in pdf.pages:
    try:
        lenn = len(i.extract_text())
        print('Page',i.page_number ,':' ,len(i.extract_text()))
    except TypeError:
        lenn = 0
        print('Page',i.page_number,':',0)
    if lenn == 0:
        print(i.page_number)
    else:
        leng += 1
        tables.append(i.extract_tables())
        print(i.extract_tables())
    text = i.extract_text()
    try:
        vkn =  re.findall(r"VKN:\s+(\d+)",text)
        vkn1 = ['VKN-1']
        vkn2 = ['VKN-2']
        vkn1.append(vkn[0])
        vkn2.append(vkn[1])
        VKN1 = pd.DataFrame(vkn1)
        VKN2 = pd.DataFrame(vkn2)
    except TypeError:
        print('scanned pdf ')
        continue
    tb_u1 = []
    tb_u2 = []
    tb_u = []
    
    for tb in range(len(tables[leng-1][0])):
        tb_u1.append(tables[leng-1][0][tb][1])
        tb_u2.append(tables[leng-1][0][tb][0])
    tb_u.append(tb_u2)
    tb_u.append(tb_u1)
        
    middle_table = []
    bottom_table = []
    for row in tables[leng-1][1]:
        non_null_row = [cell for cell in row if cell is not None]
        
        if len(non_null_row) > 2:
            middle_table.append(non_null_row)
        else:
            bottom_table.append(non_null_row)

    b_u1 = []
    b_u2 = []
    b_u = []
    for i in range(len(bottom_table)):
        b_u1.append(bottom_table[i][1])
        b_u2.append(bottom_table[i][0])
    b_u.append(b_u2)
    b_u.append(b_u1)

    u = pd.DataFrame(tb_u)
    o = pd.DataFrame(middle_table)
    a = pd.DataFrame(b_u)

    df = pd.DataFrame(pd.concat([VKN1,VKN2,u,o,a],axis=1,ignore_index = True))
    fat = fat.append(df)
    

for i, val in enumerate(fat.columns.values):
        fat.columns.values[i] = (str(i))



   
root = tkinter.Tk()
root.withdraw()
root.update()
path_save = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=(("Excel workbook", "*.xlsx"), ('Comma separated value','.csv') ,("All Files", "*.*")))
root.destroy()

import xlsxwriter

if path_save.split('.')[1] == 'xlsx':
    writer = pd.ExcelWriter(path_save.split('.')[0]+'.xlsx', engine='xlsxwriter')
    fat.to_excel(writer)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    (max_row, max_col) = fat.shape

    # Create a list of column headers, to use in add_table().
    column_settings = [{'header': str(column)} for column in fat.columns]

    # Add the Excel table structure. Pandas will add the data.
    worksheet.add_table(0, 0, max_row, max_col , {'columns': column_settings})

    # Make the columns wider for clarity.
    worksheet.set_column(0, max_col - 1, 12)
    border_fmt = workbook.add_format({'text_wrap': True,'bottom':1, 'top':1, 'left':1, 'right':1})
    worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(fat), len(fat.columns)), {'type': 'no_errors', 'format': border_fmt})

    writer.save()

    
elif path_save.split('.')[1] == 'csv':
    fat.to_excel(path_save+'.csv')
else:
    print('Excel uzantısı seçiniz')



