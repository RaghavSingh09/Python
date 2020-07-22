import csv
import os
import xlrd
import xlwt
import tkinter as tk
import webbrowser
from tkinter import messagebox,filedialog
from xlwt import Workbook
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror

window = tk.Tk()
window.title("Python File Row Count App")
window.resizable(0, 0)

lblInFileName = tk.Label(window, text="Enter Filename or Folder Location:", font=("Arial Bold", 10))
lblExcelSheetInfo = tk.Label(window, text="Enter Excel Sheet Number \ Name:", font=("Arial Bold", 10))
lblOutFileName = tk.Label(window, text="Enter Output File Location:             ", font=("Arial Bold", 10))

lblInFileName.grid (column=0, row=2)
lblExcelSheetInfo.grid (column=0, row=4)
lblOutFileName.grid (column=0, row=6)

ipFileText = tk.StringVar()
opFileText = tk.StringVar()
inFileName = tk.Entry(window,width=35,textvariable=ipFileText)
excelSheetInfo = tk.Entry(window,width=35)
outFileName = tk.Entry(window,width=35,textvariable=opFileText)

inFileName.grid(column=1, row=2)
excelSheetInfo.grid(column=1, row=4)
outFileName.grid(column=1, row=6)

var = tk.IntVar()
var.set(1)
def sel():
    rop=str(var.get())
    return rop

rdnFLF = tk.Radiobutton(window, text='Folder List In Excel', variable=var, value=1, command=sel)
rdnSFolder = tk.Radiobutton(window, text='Single Folder         ', variable=var, value=2, command=sel)
rdnSFile = tk.Radiobutton(window, text='Single File              ', variable=var, value=3, command=sel)
rdnFLF.grid(column=1, row=8)
rdnSFolder.grid(column=1, row=10)
rdnSFile.grid(column=1, row=12)

def load_Infile():
    fname = askopenfilename(filetypes=(("All files", "*.*"),("Excel files", "*.xlsx"),("CSV files", "*.csv"),("SQL files", "*.sql"),("Text files", "*.txt"),))
    if fname:
        try:
            ipFileText.set(fname)
        except:
            showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        return
def load_Outfile():
    fname = askopenfilename(filetypes=(("Excel files", "*.xls"),("CSV files", "*.csv"),))
    if fname:
        try:
            opFileText.set(fname)
        except:
            showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        return
def openHelp(url):
    #webbrowser.open_new(url)
    help='''***************************************************************************
Mandatory Requirements:-
1. Folder List In Excel
    a. Enter excel file name in which you are storing the file or folder info
    b. Enter the worksheet name or number
    c. Enter output file location
2. Single Folder
    a. Enter the folder location to get files row count in it
    b. Enter output file location
3. Single File
    a. Enter the file location with name to get its row count
    b. If excel file is given then enter the worksheet name or number
        
***************************************************************************
    '''
    messagebox.showinfo('Help Message',help )
def openFile(url):
    webbrowser.open_new(url)

def fn_FileRowCount(fileName,sheetInfo):
    if fileName.find(".xls")==-1:
        with open(fileName,"r") as fr:
            reader = csv.reader(fr,delimiter = ",")
            data = list(reader)
            row_count = len(data)
            return fileName,str(row_count)
    else:
        wb = xlrd.open_workbook(fileName)
        try:
            sheet = wb.sheet_by_name(sheetInfo)
        except:
            sheet = wb.sheet_by_index(int(sheetInfo)-1)
        row_count=sheet.nrows
        return fileName,str(row_count)

def clicked():

    rdBtn = sel()
    ifile = inFileName.get()
    ofile = outFileName.get()
    xcelInfo = excelSheetInfo.get()

    if(xcelInfo==''):
        xcelInfo=1
    try:
        if((ifile=='' or ofile=='') and (rdBtn ==1 or rdBtn==2)):
            messagebox.showinfo('Error', 'Please Provide Input / Output File Info.')
        elif(ifile!='' and ofile!='' and rdBtn=='1'):
            #Reading Folder Loactions From Excel File
            inputWB = xlrd.open_workbook(ifile)
            try:
                excelSheet = inputWB.sheet_by_index(int(xcelInfo)-1)
            except:
                excelSheet = inputWB.sheet_by_name(xcelInfo)
            folderCnt=excelSheet.nrows
            #Creating output Excel File
            outputWB = Workbook()
            style = xlwt.easyxf('font: bold 1, color blue;')
            outputSheet = outputWB.add_sheet('RowCountOutput')

            outputSheet.write(0,0,"Files with folder location",style)
            outputSheet.write(0,1,"Row Count",style)

            outputRc = 1
            fileLoc=""
            for r in range(1,folderCnt):
                value=excelSheet.row_values(r)
                fileLoc = fileLoc.join(value)
                if os.path.isfile(fileLoc):
                    fn, rc =  fn_FileRowCount(fileLoc,1)
                    outputSheet.write(outputRc,0,fn)
                    outputSheet.write(outputRc,1,rc)
                    outputRc += 1
                elif os.path.isdir(fileLoc):
                    for root, directories, files in os.walk(fileLoc, topdown=False):
                        for name in files:
                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                            outputSheet.write(outputRc,0,fn)
                            outputSheet.write(outputRc,1,rc)
                            outputRc += 1
                else:
                    print("File or Directory Does Not Exists")

            outputWB.save(ofile)
            messagebox.showinfo('Message', "Row Count Fetched Successfully!!!")
            openFilelink.config(text='Open File')
            #webbrowser.open_new(ofile)
        elif(ifile!='' and ofile!='' and rdBtn=='2'):
            fileLoc = ifile
            outputWB = Workbook()
            style = xlwt.easyxf('font: bold 1, color blue;')
            outputSheet = outputWB.add_sheet('RowCountOutput')

            outputSheet.write(0,0,"Files with folder location",style)
            outputSheet.write(0,1,"Row Count",style)

            outputRc = 1
            if os.path.isdir(fileLoc):
                for root, directories, files in os.walk(fileLoc, topdown=False):
                    for name in files:
                        fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                        outputSheet.write(outputRc,0,fn)
                        outputSheet.write(outputRc,1,rc)
                        outputRc += 1
                outputWB.save(ofile)
                messagebox.showinfo('Message', "Row Count Fetched Successfully!!!")
                openFilelink.config(text='Open File')
                #webbrowser.open_new(ofile)
            else:
                messagebox.showinfo('Error', "Not Valid Folder Location!")
        elif(ifile!='' and rdBtn=='3'):
            if os.path.isfile(ifile):
                fn, rc =  fn_FileRowCount(ifile,xcelInfo)
                messagebox.showinfo('Message', fn+" has "+rc+" rows !")
            else:
                messagebox.showinfo('Error', "Not A Valid File!")
    except Exception as e:
        messagebox.showinfo('Error', 'Someting Went Wrong ! \n'+str(e))

btnIF = tk.Button(window, text="...", command=load_Infile, width=2,height = 1)
btnIF.grid(column=2,row=2)
btnOF = tk.Button(window, text="...", command=load_Outfile, width=2,height = 1)
btnOF.grid(column=2,row=6)
helplink1 = tk.Label(window, text="Help", fg="blue", cursor="hand2")
helplink1.grid (column=3, row=2)
helplink1.bind("<Button-1>", lambda e: openHelp("http://www.google.com"))
openFilelink = tk.Label(window, text="", fg="blue", cursor="hand2")
openFilelink.grid (column=0, row=12)
openFilelink.bind("<Button-1>", lambda e: openFile(outFileName.get()))
btnGRC = tk.Button(window, text="Get Row Count",bg="red", fg="black",width = 12, height = 2, bd = 5,command=clicked, cursor="hand2")
btnGRC.grid (column=0, row=10)

window.geometry('500x200')
window.mainloop()