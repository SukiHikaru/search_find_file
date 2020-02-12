import tkinter as TK
from tkinter import messagebox
from tkinter import filedialog
import openpyxl as op
import xml.etree.cElementTree as et
import pprint as pp

root = TK.Tk()
root.geometry('300x300')
root.iconbitmap(r'cute.ico')
root.title('Sort Data')

cell_content = {}
str_tag = '{http://www.tracetronic.de/xml/ecu-test}'

bus_key = str_tag + 'BUS-KEY'
node_name = str_tag + 'NODE-NAME'
frame_name = str_tag + 'FRAME-NAME'
pdu_name = str_tag + 'PDU-NAME'
signal_name = str_tag + 'SIGNAL-NAME'
# user could open browser for choosing file

def browse_file():
    path = TK.filedialog.askopenfilename()  # get input from user
    return path

def getDataExcel():
    global cell_content   # store place for excel data
    cell_content = {}
    global list_excel
    list_excel = []
    xlsx_string_path = browse_file()  # get user path from browse_file function
    wb = op.load_workbook(xlsx_string_path)  # load workbook
    sheet = wb['Tabelle1']
    for i in range(1, sheet.max_row + 1):
        cell_content[sheet['A' + str(i)].value] = xlsx_string_path
    for keys in cell_content:
        list_excel.append(keys)
    #pp.pprint(list_excel)


def getDataPgk():
    # BUS-KEY/NODE-NAME/FRAME-NAME/PDU-NAME/SIGNAL-NAME'}

    tree = et.parse(r'C:\Users\HoNg1\Documents\Internship work\Tooling\Bus_Finder\Test_1\StatusCentralLockingSystem.pkg')
    root = tree.getroot()  # get the root of the pkg file

    for mapping in root:
        for mapping_item in mapping:
            for xaccess in mapping_item.iter(str_tag + 'XACCESS'):
                for i in xaccess.attrib:
                    if xaccess.attrib[i] == 'xaBusSignalVariable':
                        global test_dict
                        test_dict = {}
                        global list_pgk
                        list_pgk = []
                        for elements in xaccess:
                            if elements.tag == bus_key:
                                test_dict[elements.tag.replace(str_tag, '')] = elements.text
                                list_pgk.insert(0,elements.text)
                            if elements.tag == node_name:
                                test_dict[elements.tag.replace(str_tag, '')] = elements.text
                                list_pgk.insert(1,elements.text)
                            if elements.tag == frame_name:
                                test_dict[elements.tag.replace(str_tag, '')] = elements.text
                                list_pgk.insert(2,elements.text)
                            if elements.tag == pdu_name:
                                test_dict[elements.tag.replace(str_tag, '')] = elements.text
                                list_pgk.insert(3,elements.text)
                            if elements.tag == signal_name:
                                test_dict[elements.tag.replace(str_tag, '')] = elements.text
                                list_pgk.insert(4,elements.text)
                        list_pgk = '/'.join(list_pgk)
                        #print(test_dict)
                        #print(list_pgk)


def compare():
    for i in range(len(list_excel)):
        if list_excel[i] == list_pgk:




label_excel = TK.Label(root, text='Excel')
label_excel.grid(row=0, column=0)

btn_excel = TK.Button(root, text='Browse', command=getDataExcel)
btn_excel.grid(row=0, column=1)

label_pkg = TK.Label(root, text='pkg')
label_pkg.grid(row=1, column=0)

btn_pkg = TK.Button(root, text='Browse', command=getDataPgk)
btn_pkg.grid(row=1, column=1)

btn_test = TK.Button(root, text='Test', command=compare)
btn_test.grid(row=2, column=1)


root.mainloop()
