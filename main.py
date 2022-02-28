from tkinter import *
from tkinter import filedialog
import os
import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl.styles import Alignment, Border, Side

root = Tk()
root.title('Sasol De Bron - GP Analyzer ')
root.geometry('350x225')

def gp_analysis():
    # Get Less and Greater than amounts
    less = int(less_than_input.get())
    greater = int(greater_than_input.get())

    less_than_input.delete(0,END)
    greater_than_input.delete(0,END)

    # Read in excell and convert to list
    check_option = c.get()
    if check_option == 0:
        file = 'Analysis Info/Analysis Info.xls'
    else:
        path = os.getcwd()
        file = filedialog.askopenfilename(initialdir=path, title='Select A File', filetypes=(("Excel 97-2003", '*.xls'),('Excel file',"*.xlsx")))
    
    read_info = pd.read_excel(file, header=1)
    data = read_info.to_numpy().tolist()

    # Remove nan and unwated data from data list
    simple_list = []
    for x in data:
        if x[0] == float("nan") or x == '----------------------------':
            pass
        elif type(x[0]) == int:
            simple_list.append(x)

    printing_list = []
    for x in simple_list:
        info.config(text='Checking')
        # print(f'Checking if {x[1]} GP {x[-1]} <=-5 or >=5')
        if x[-1] <=-less or x[-1] >=greater:
            printing_list.append(x)

    items = len(printing_list) + 2

    wb = Workbook()
    ws = wb.active
    ws.title = 'Stock GP List'
    # print('Create Excel Workbook')

    # Set Column Width 
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 35

    # Set row names
    # print('Create Column Names')
    ws['C1'] = 'Pack'
    ws['E1'] = 'Desired %'
    ws['F1'] = 'Sugg Incl'
    ws['G1'] = 'Selling A'
    ws['H1'] = 'Var Amt on'
    ws['I1'] = 'GP % on'
    ws['J1'] = 'Var % on'

    ws['A2'] = 'Pack Code'
    ws['B2'] = 'Pack Description'
    ws['C2'] = 'Size'
    ws['D2'] = 'Sys Cost'
    ws['E2'] = 'on Sell A'
    ws['F2'] = 'Selling A'
    ws['G2'] = 'Inclusive'
    ws['H2'] = 'Sell A Incl'
    ws['I2'] = 'Sys Cost'
    ws['J2'] = 'Sys Cost'

    # # loop through sheet totals
    i = 3
    while i <= items:
        for stock in printing_list:
            ws[f'A{i}'] = stock[0]
            ws[f'B{i}'] = stock[1]
            ws[f'C{i}'] = stock[2]
            ws[f'D{i}'] = stock[3]
            ws[f'E{i}'] = stock[4]
            ws[f'F{i}'] = stock[5]
            ws[f'G{i}'] = stock[6]
            ws[f'H{i}'] = stock[7]
            ws[f'I{i}'] = stock[8]
            ws[f'J{i}'] = stock[9]
            
            ws[f'A{i}'].number_format = '0'
            i += 1
    info.config(text='Analysis Complete')
    f_name = file_name_input.get()
    if f_name == '':
        wb.save(f'GP Analysis Less than {less} and Greater Than {greater}.xlsx')
    else:
        wb.save(f'{f_name}.xlsx')
        file_name_input.delete(0,END)

less_than = Label(root, text='Less Than (-)')
less_than.grid(row=0, column=0, padx=(10,0), pady=(10,0))

less_than_input = Entry(root, text='Less Than (-)')
less_than_input.grid(row=0, column=1, padx=(10,0), pady=(10,0))

greater_than = Label(root, text='Greater Than')
greater_than.grid(row=1, column=0, padx=(10,0), pady=(10,0))

greater_than_input = Entry(root, text='Greater Than')
greater_than_input.grid(row=1, column=1, padx=(10,0), pady=(10,0))

file_name = Label(root, text='File Name')
file_name.grid(row=2, column=0, padx=(10,0), pady=(10,0))

file_name_input = Entry(root, text='File Name')
file_name_input.grid(row=2, column=1, padx=(10,0), pady=(10,0))

c = IntVar()
check = Checkbutton(root, text='I will Select File', variable=c)
check.grid(row=3, column=0, padx=(10,0), pady=(10,0))

cal = Button(root, text='Start Analysis', command=gp_analysis)
cal.grid(row=4, column=0, padx=(10,0), pady=(10,0))

info = Label(root, text='')
info.grid(row=5, column=0)

# Run program
root.mainloop()



