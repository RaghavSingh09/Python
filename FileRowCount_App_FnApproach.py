import csv
import os
import xlrd
import xlwt
from xlwt import Workbook
import webbrowser
import datetime
try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:
    import Tkinter as tk
    import ttk
from tkinter import messagebox,filedialog
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
from tkcalendar import Calendar, DateEntry
#import tkcalendar

window = tk.Tk()
window.title("Python File Row Count App")
window.resizable(0, 0)
calDate = ''
lblInFileName = tk.Label(window, text="Enter Filename or Folder Location:", font=("Arial Bold", 10))
lblExcelSheetInfo = tk.Label(window, text="Enter Excel Sheet Number \ Name:", font=("Arial Bold", 10))
lblOutFileName = tk.Label(window, text="Enter Output File Location:", font=("Arial Bold", 10))
lblFileDate = tk.Label(window, text="Enter File Modified Date (MM/DD/YYYY):", font=("Arial Bold", 10))

lblInFileName.grid (column=0, row=2,sticky = "W")
lblExcelSheetInfo.grid (column=0, row=4,sticky = "W")
lblOutFileName.grid (column=0, row=6,sticky = "W")
lblFileDate.grid (column=0, row=8,sticky = "W")

ipFileText = tk.StringVar()
opFileText = tk.StringVar()
fileCompDate= tk.StringVar()

inFileName = tk.Entry(window,width=35,textvariable=ipFileText)
excelSheetInfo = tk.Entry(window,width=35)
outFileName = tk.Entry(window,width=35,textvariable=opFileText)
fileDate = tk.Entry(window,width=15,textvariable=fileCompDate)

inFileName.grid(column=1, row=2)
excelSheetInfo.grid(column=1, row=4)
outFileName.grid(column=1, row=6)
fileDate.grid(column=1, row=8,sticky = "W")

var = tk.IntVar()
var.set(1)
def sel():
    rop=str(var.get())
    return rop

rdnFLF = tk.Radiobutton(window, text='Folder List From File', variable=var, value=1, command=sel)
rdnSFolder = tk.Radiobutton(window, text='Single Folder', variable=var, value=2, command=sel)
rdnSFile = tk.Radiobutton(window, text='Single File', variable=var, value=3, command=sel)
rdnFLF.grid(column=1, row=10,sticky = "W")
rdnSFolder.grid(column=1, row=12,sticky = "W")
rdnSFile.grid(column=1, row=14,sticky = "W")

def load_Infile():
    if(sel()=='2'):
        fname = filedialog.askdirectory()
    else:
        fname = askopenfilename(filetypes=(("All files", "*.*"),("Excel files", "*.xlsx"),("CSV files", "*.csv"),("SQL files", "*.sql"),("Text files", "*.txt")))
    if fname:
        try:
            ipFileText.set(fname)
        except:
            showerror("Open Source File/Folder", "Failed to read file/folder\n'%s'" % fname)
        return

def load_Outfile():
    fname = askopenfilename(filetypes=(("Excel files", "*.xls*"),("CSV files", "*.csv"),("Text files", "*.txt")))
    if fname:
        try:
            opFileText.set(fname)
        except:
            showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        return

def openHelp(url):
    #webbrowser.open_new(url)
    help='''***************************************************************************
Mandatory Requirements:-\n
1. Folder List From File
    A. Enter file name in which you are storing the file or folder info
    B. Enter the worksheet name or number if excel file is given
    C. Enter output file location\n
2. Single Folder
    A. Enter the folder location to get files row count in it
    B. Enter output file location\n
3. Single File
    A. Enter the file location with name to get its row count
    B. If excel file is given then enter the worksheet name or number\n
4.  Date
    A. Enter date to get row count of specific file with last modified date
    B. You can either input date manually or by using calendar button
    C. If date field is empty you will get row count of all files in the folder\n
Exceptions:-
    A. For Folder List In File option input file can be excel, csv or text file
    B. Output file can be saved only in xlsx, xls, csv or txt format
    C. You can get row count on only 1st sheet of excel files in a folder
        
***************************************************************************
    '''
    messagebox.showinfo('Help Message',help)

def calendarDate():
    def getDate_Sel(e):
        fileCompDate.set(cal.get_date())

    cal = DateEntry(window, width=12, background='darkblue', foreground='white', borderwidth=2)
    cal.grid(column=1, row=8,sticky = "W")
    cal.bind("<<DateEntrySelected>>", getDate_Sel)

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

'''
def fn_OutputFileType(ofile,counter,fn,rc):
    if(ofile.find('.csv')>1 or ofile.find('.txt')>1):
        if(counter==1):
            with open(ofile, mode='w') as output_file:
                fieldnames = ['File Names', 'Row Count']
                output_writer = csv.DictWriter(output_file, fieldnames=fieldnames, delimiter=',')
                output_writer.writeheader()
        else:
                output_writer.writerow({'File Names' : fn,'Row Count' : rc})
    elif(ofile.find('.xls')>1):
        if(counter==1):
            outputWB = Workbook()
            style = xlwt.easyxf('font: bold 1, color blue;')
            outputSheet = outputWB.add_sheet('RowCountOutput')

            outputSheet.write(0,0,"Files with folder location",style)
            outputSheet.write(0,1,"Row Count",style)
        else:
            outputSheet.write(counter,0,fn)
            outputSheet.write(counter,1,rc)
'''

def fn_FileSave(ifile,ofile,xcelInfo,fDate=''):
    if ifile.find(".xls")==-1:
        with open(ofile, mode='w', newline='') as output_file:
            fieldnames = ['File Names', 'Row Count']
            output_writer = csv.DictWriter(output_file, fieldnames=fieldnames, delimiter=',')
            output_writer.writeheader()
            if os.path.isdir(ifile):
                for root, directories, files in os.walk(ifile, topdown=False):
                    for name in files:
                        if(fDate==''):
                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                            output_writer.writerow({'File Names' : fn,'Row Count' : rc})
                        elif(datetime.date.fromtimestamp(os.path.getmtime(ifile+r'\\'+name)).strftime("%m/%d/%Y") == fDate):
                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                            output_writer.writerow({'File Names' : fn,'Row Count' : rc})
            else:
                with open(ifile,"r") as csv_file:
                    csv_reader = csv.reader(csv_file)
                    line_count = 1
                    fileLoc = ''
                    for row in csv_reader:
                        if line_count == 1:
                            line_count += 1
                        else:
                            fileLoc = fileLoc.join(row)
                            if os.path.isfile(fileLoc):
                                if(fDate==''):
                                    fn, rc =  fn_FileRowCount(fileLoc,1)
                                    output_writer.writerow({'File Names' : fn,'Row Count' : rc})
                                elif(datetime.date.fromtimestamp(os.path.getmtime(fileLoc)).strftime("%m/%d/%Y") == fDate):
                                    fn, rc =  fn_FileRowCount(fileLoc,1)
                                    output_writer.writerow({'File Names' : fn,'Row Count' : rc})
                            elif os.path.isdir(fileLoc):
                                for root, directories, files in os.walk(fileLoc, topdown=False):
                                    for name in files:
                                        if(fDate==''):
                                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                                            output_writer.writerow({'File Names' : fn,'Row Count' : rc})
                                        elif(datetime.date.fromtimestamp(os.path.getmtime(fileLoc+r'\\'+name)).strftime("%m/%d/%Y") == fDate):
                                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                                            output_writer.writerow({'File Names' : fn,'Row Count' : rc})
        output_file.close()
    else:
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
                if(fDate==''):
                    fn, rc =  fn_FileRowCount(fileLoc,1)
                    outputSheet.write(outputRc,0,fn)
                    outputSheet.write(outputRc,1,rc)
                    outputRc += 1
                elif(datetime.date.fromtimestamp(os.path.getmtime(fileLoc)).strftime("%m/%d/%Y") == fDate):
                    fn, rc =  fn_FileRowCount(fileLoc,1)
                    outputSheet.write(outputRc,0,fn)
                    outputSheet.write(outputRc,1,rc)
                    outputRc += 1
            elif os.path.isdir(fileLoc):
                for root, directories, files in os.walk(fileLoc, topdown=False):
                    for name in files:
                        if(fDate==''):
                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                            outputSheet.write(outputRc,0,fn)
                            outputSheet.write(outputRc,1,rc)
                            outputRc += 1
                        elif(datetime.date.fromtimestamp(os.path.getmtime(fileLoc+r'\\'+name)).strftime("%m/%d/%Y") == fDate):
                            fn, rc =  fn_FileRowCount(os.path.join(root, name),1)
                            outputSheet.write(outputRc,0,fn)
                            outputSheet.write(outputRc,1,rc)
                            outputRc += 1
            else:
                print("File or Directory Does Not Exists")
        outputWB.save(ofile)

def clicked():

    rdBtn = sel()
    ifile = inFileName.get()
    ofile = outFileName.get()
    xcelInfo = excelSheetInfo.get()
    if(fileDate.get()!=''):
        fDate = datetime.date.strftime(datetime.datetime.strptime(fileDate.get(),"%Y-%m-%d"),"%m/%d/%Y")
    else:
        fDate = ''

    if(xcelInfo==''):
        xcelInfo=1

    if(fDate==''):
        try:
            if((ifile=='' or ofile=='') and (rdBtn !='3')):
                messagebox.showinfo('Error', 'Please Provide Input / Output File Info.')
            elif(ifile!='' and ofile!='' and rdBtn=='1'):
                fn_FileSave(ifile,ofile,xcelInfo)
                messagebox.showinfo('Message', "Row Count Fetched Successfully!!!")
                openFilelink.config(text='Open File')
            elif(ifile!='' and ofile!='' and rdBtn=='2'):
                fn_FileSave(ifile,ofile,xcelInfo)
                messagebox.showinfo('Message', "Row Count Fetched Successfully!!!")
                openFilelink.config(text='Open File')
            elif(ifile!='' and rdBtn=='3'):
                if os.path.isfile(ifile):
                    fn, rc =  fn_FileRowCount(ifile,xcelInfo)
                    messagebox.showinfo('Message',"File Name : " +fn+"\nRow Count : "+rc)
                else:
                    messagebox.showinfo('Error', "Not A Valid File!")
        except Exception as e:
            messagebox.showinfo('Error', 'Someting Went Wrong ! \n'+str(e))
    else:
        try:
            if((ifile=='' or ofile=='') and (rdBtn !='3')):
                messagebox.showinfo('Error', 'Please Provide Input / Output File Info.')
            elif(ifile!='' and ofile!='' and rdBtn=='1'):
                fn_FileSave(ifile,ofile,xcelInfo,fDate)
                messagebox.showinfo('Message', "Row Count Fetched Successfully!!!")
                openFilelink.config(text='Open File')
            elif(ifile!='' and ofile!='' and rdBtn=='2'):
                fn_FileSave(ifile,ofile,xcelInfo,fDate)
                messagebox.showinfo('Message', "Row Count Fetched Successfully!!!")
                openFilelink.config(text='Open File')
            elif(ifile!='' and rdBtn=='3'):
                if os.path.isfile(ifile):
                    fn, rc =  fn_FileRowCount(ifile,xcelInfo)
                    messagebox.showinfo('Message',"File Name : " +fn+"\nRow Count : "+rc)
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
openFilelink.grid (column=0, row=14)
openFilelink.bind("<Button-1>", lambda e: openFile(outFileName.get()))
btnGRC = tk.Button(window, text="Get Row Count",bg="red", fg="black",width = 12, height = 2, bd = 5,command=clicked, cursor="hand2")
btnGRC.grid (column=0, row=12)
btnCal = ttk.Button(window, text='Calendar',command=calendarDate, cursor="hand2")
btnCal.grid (column=1, row=8,sticky = "E")

window.geometry('525x220')
window.mainloop()