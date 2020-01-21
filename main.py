import tkinter as TK
import openpyxl as op
from tkinter import *
from tkinter import filedialog
import xml.etree.cElementTree as et

root = TK.Tk()
root.geometry('100x100')


def browse_file():
    path = TK.filedialog.askopenfilename()
    return path


def getDataExcel():
    xlsx_string_path = browse_file()  # get user path from browse_file function
    wb = op.load_workbook(xlsx_string_path)  # load workbook
    sheet = wb['Tabelle1']
    for i in range(1, sheet.max_row + 1):
        cell_content = sheet['A' + str(i)].value
        if cell_content != None:
            print(cell_content)


def getDataPgk():
    pkg_string_path = browse_file()
    tree = et.parse(pkg_string_path)  # load path into tree
    root = tree.getroot()  # get the start of the pkg file
    print(root.tag)


label_excel = TK.Label(root, text='Excel')
label_excel.grid(row=0, column=0)

btn_excel = TK.Button(root, text='Browse', command=getDataExcel)
btn_excel.grid(row=0, column=1)

label_pkg = TK.Label(root, text='pkg')
label_pkg.grid(row=1, column=0)

btn_pkg = TK.Button(root, text='Browse', command=getDataPgk)
btn_pkg.grid(row=1, column=1)

btn_test = TK.Button(root, text='Test')
btn_test.grid(row=2, column=1)

root.mainloop()
