import tkinter as tk
from tkinter import messagebox,filedialog
import pyodbc as pdb
import csv
import os
import xlrd
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import re

window = tk.Tk()
window.title("Python GY-PS Data Load UI")
window.resizable(0, 0)

lblFileName = tk.Label(window, text="Enter Filename With Folder Location:", font=("Arial", 10))
lblServerAdd = tk.Label(window, text="Enter SQL Server Address:             ", font=("Arial", 10))
lblDBName = tk.Label(window, text="Enter Database Name To Connect:  ", font=("Arial", 10))
lblUN = tk.Label(window, text="Enter SQL Server Username:          ", font=("Arial", 10))
lblPWD = tk.Label(window, text="Enter SQL Server Password:           ", font=("Arial", 10))
lblCommand = tk.Label(window, text="Enter Table Name \ SQL Command:", font=("Arial", 10))
lblFileLoadType = tk.Label(window, text="File Load Type:                ", font=("Arial", 10))
lblMsg = tk.Label(window, text="", font=("Arial Bold", 10),fg='blue')
lblSNCV = tk.Label(window, text="SQL Native Client Version: ", font=("Arial", 10))
lblCounter = tk.Label(window,text='',font=("Arial", 10),fg="green")
lblExcelInfo = tk.Label(window, text="Enter Excel Sheet Number \ Name:  ", font=("Arial", 10))

lblFileName.grid (column=0, row=6)
lblServerAdd.grid (column=0, row=10)
lblDBName.grid (column=0, row=12)
lblUN.grid (column=0, row=14)
lblPWD.grid (column=0, row=16)
lblCommand.grid (column=0, row=18)
lblSNCV.grid (column=0, row=20)
lblFileLoadType.grid (column=0, row=22)
lblMsg.grid (column=0, row=24)
lblCounter.grid(column=3,row=28)
lblExcelInfo.grid (column=0, row=8)

entryText = tk.StringVar()
varID = tk.IntVar()
varRS = tk.IntVar()
varCT = tk.IntVar()
varTT = tk.IntVar()
varTT.set(1)
var = tk.IntVar()

filename = tk.Entry(window,width=30,textvariable=entryText )
filename.grid(column=3, row=6)
sheetinfo = tk.Entry(window,width=30)
sheetinfo.grid(column=3, row=8)
serverName = tk.Entry(window,width=30)
serverName.grid(column=3, row=10)
dbName = tk.Entry(window,width=30)
dbName.grid(column=3, row=12)
userName = tk.Entry(window,width=30)
userName.grid(column=3, row=14)
password = tk.Entry(window,width=30)
password.grid(column=3, row=16)
tableName = tk.Text(window, font=("Arial", 9),width=26,height=4)
tableName.grid(column=3, row=18)
nativeCV = tk.Spinbox(window, from_ = 11, to = 20,width=10)
nativeCV.grid(column=3,row=20)

def sel():
    rop=str(var.get())
    return rop

def varCT_states():
    ct=str(varCT.get())
    return ct

def varTT_states():
    tt=str(varTT.get())
    return tt

def varRS_states():
    rs=str(varRS.get())
    return rs

def varID_states():
    id=str(varID.get())
    return id

def load_file():
    fname = askopenfilename(filetypes=(("All files", "*.*"),
                                       ("Excel files", "*.xlsx"),
                                       ("CSV files", "*.csv"),
                                       ("SQL files", "*.sql"),
                                       ("Text files", "*.txt"),
                                       ))
    if fname:
        try:
            entryText.set(fname)
            #filename.config(textvariable=entryText,fg="red")
        except:
            showerror("Open Source File", "Failed to read file\n'%s'" % fname)
        return

def callback(url):
    #webbrowser.open_new(url)
    #webbrowser.open_new(r"file://c:\test\test.csv")
    #help='TT - Truncate Table\nCT - Create Table\nRS - Run Script\nNon-TDC - Normal CSV or Excel File To Load Data\nTDC - GoodYear Template TDC File in CSV Format\nAttachment - Load Attachment Files From Folder\n... - Browse Files To Load Data\n\n***Conditions***\n\n'
    help='''***Check Boxes***\nID - Import Data To CSV\nTT - Truncate Table\nCT - Create Table\nRS - Run Script\n\n***Radio Buttons***
Non-TDC - Normal CSV or Excel File\nTDC - GoodYear Template TDC File in CSV Format\nAttachment - Load attachment files from a folder\n... - Browse Files To Load Data\n
***Conditions***\n
1. If loading data from excel file then mention workbook/sheet number/name.\n
2. If you want run a SQL command then enter it in SQL command text box and select RS check box.\n
3. If you are using windows authenticaton then SQL UN and PWD is not required.\n
***Exceptions***\n
1. CSV files will not be able to load multi-language data, use excel file.\n
2. If excel file contains date column then please update those columns in your table to get correct date.
   DATEADD(Day,CAST(Your_Col_Name AS INT)-2,'1900-01-01')\n
3. If you are running a SQL script file then make sure it contains Go after each SQL block and select RS check box.\n
'''
    messagebox.showinfo('Help Message',help )

tdcType = tk.Radiobutton(window, text='TDC File      ', variable=var, value=2, command=sel)
nonTdcType = tk.Radiobutton(window, text='Non TDC    ', variable=var, value=1, command=sel)
loadAttach = tk.Radiobutton(window, text='Attachment', variable=var, value=3, command=sel)
createTable = tk.Checkbutton(window, text='CT', variable=varCT)
truncateTable = tk.Checkbutton(window, text='TT', variable=varTT)
runScript = tk.Checkbutton(window, text='RS', variable=varRS)
importData = tk.Checkbutton(window, text='ID', variable=varID)

nonTdcType.grid(column=3, row=22)
tdcType.grid(column=3, row=23)
loadAttach.grid(column=3, row=24)
truncateTable.grid(column=4, row=22)
createTable.grid(column=4, row=23)
runScript.grid(column=4, row=24)
importData.grid(column=4, row=20)

def get_spinbox_value():
    sbv=str(nativeCV.get())
    return sbv

def clicked():

    filename1 = filename.get()
    sheetInfo1=sheetinfo.get()
    serverName1 = serverName.get()
    dbName1 = dbName.get()
    userName1 = userName.get()
    password1 = password.get()
    tableName1 = tableName.get("1.0",'end-1c')
    radioBtn = sel()
    spinBox = get_spinbox_value()
    lblMsg.config(text='')
    truncateTable1=varTT_states()
    createTable1=varCT_states()
    runScript1=varRS_states()
    importData1=varID_states()

    '''
    filename1 = '09_DM-AI-C__Project_Custom_Fields (R6).xlsx'
    serverName1 = 'IN5CD7031HD8\SSMS_TEST'
    dbName1 = 'PMT_DM'
    userName1 = 'gbrs_dp'
    password1 = 'gbrs_dp'
    tableName1 = 'Temp_DMAIC_Custom_field'
    radioBtn = '1'
    spinBox = '11'
    '''

    if (radioBtn=='0' and runScript1=='0' and importData1=='0'):
        #lblMsg.config(text='Error Msg:\n' +'Please Select File Load Type !',fg="red")
        messagebox.showinfo('ERROR', 'Please Select File Load Type !')

    if (radioBtn=='0' and importData1=='1' and filename1!=''):
        try:

            if(userName1!='' and password1!=''):
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
            else:
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

            cursor = conn.cursor()
            #sql = """exec sp_codesearch 'GBS'"""
            cursor.execute(tableName1)
            if(filename1.endswith(".csv")):
                with open(filename1, "w", newline='') as f:
                    wrtr = csv.writer(f)
                    wrtr.writerow([i[0] for i in cursor.description]) # write headers
                    for row in cursor:
                        wrtr.writerow(row)

                messagebox.showinfo('Message', 'Data Imported Successfully')
            else:
                messagebox.showinfo('Error', 'Please Select CSV File To Import Data')


        except Exception as e:
            messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))

    if (radioBtn=='0' and runScript1=='1' and filename1==''):
        try:

            if(userName1!='' and password1!=''):
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
            else:
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

            cursor = conn.cursor()
            cursor.execute(tableName1)
            cursor.commit()

            messagebox.showinfo('Message', 'Query Executed Successfully')


            if(importData1=='1'):
                #sql = """exec sp_codesearch 'GBS'"""
                cursor.execute(tableName1)
                if(filename1.endswith(".csv") or filename1.endswith(".txt")):
                    with open(filename1, "w", newline='') as f:
                        wrtr = csv.writer(f)
                        for row in cursor:
                            wrtr.writerow(row)

                    messagebox.showinfo('Message', 'Data Imported Successfully')
                else:
                    messagebox.showinfo('Error', 'Please Select CSV File To Import Data')


        except Exception as e:
            messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))


    if (radioBtn=='0' and runScript1=='1' and (filename1.endswith(".sql") or filename1.endswith(".txt"))):
        try:

            if(userName1!='' and password1!=''):
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
            else:
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

            cursor = conn.cursor()
            sqlQuery = ''

            with open(filename, 'r') as inp:
                for line in inp:
                    if line == 'GO\n':
                        cursor.execute(sqlQuery)
                        cursor.commit()
                        #print('IF  '+sqlQuery)
                        sqlQuery = ''
                    elif 'PRINT' in line:
                        disp = line.split("'")[1]
                        #print(disp, '\r')
                    else:
                        sqlQuery = sqlQuery + line
                        #print('ESLE  '+sqlQuery)

            inp.close()

            messagebox.showinfo('Message', 'Script Executed Successfully')

        except Exception as e:
            messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))


    if (radioBtn=='1'):
        try:

                if(userName1!='' and password1!=''):
                    conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
                else:
                    conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

                #with open (filename1, 'r',encoding='utf8') as f: #,encoding='utf8'
                #f = open(filename1, 'r', encoding='utf-8')
                if(filename1.endswith("csv")):
                    f=open (filename1, 'r',errors='ignore') #,encoding='utf8',errors='ignore'
                    reader = csv.reader(f)
                    headers=[]
                    data = next(reader)
                    cursor = conn.cursor()

                    if(createTable1=='1'):
                        headers=data
                        statement = "IF object_id('"+tableName1 +"', 'U') is null \nBEGIN\n CREATE TABLE "+tableName1 +"\n("
                        for i in range(len(headers)):
                            headers[i]=re.sub(r"\s+", "_", re.sub(r"[^\w\s]", '',headers[i])).rstrip('_')
                            statement = (statement + '\n{} NVARCHAR({}),').format('['+headers[i]+']', 'MAX')
                        statement = statement[:-1] + '\n)\nEND'
                        cursor.execute(statement)

                    query = 'insert into '+dbName1+'.dbo.'+tableName1+' values ({0})'
                    query = query.format(','.join('?' * len(data)))

                    if(truncateTable1=='1'):
                        cursor.execute('TRUNCATE TABLE '+dbName1+'.dbo.'+tableName1)
                    #cursor.execute(query, data)

                    counter=0
                    for data in reader:
                        cursor.execute(query, data)
                        counter += 1
                        lblCounter.config(text='Inserted Record Count: '+str(counter))

                    cursor.commit()
                    messagebox.showinfo('Message', 'Data Loaded Successfully From CSV')

                if(filename1.endswith("xlsx") or filename1.endswith("xls")):
                    loc = (filename1)
                    wb = xlrd.open_workbook(loc)
                    if(len(sheetInfo1.rstrip().lstrip())<=0):
                        sheetInfo1='1'
                    else:
                        sheetInfo1=sheetInfo1.rstrip().lstrip()

                    if (re.findall(r'^\w+',sheetInfo1) and  not re.findall(r'^\d+',sheetInfo1)):
                        sheet = wb.sheet_by_name(sheetInfo1)
                        sheet.cell_value(0, 0)
                        cursor = conn.cursor()
                        headers=[]
                        if(createTable1=='1'):
                            headers=sheet.row_values(0)
                            statement = "IF object_id('"+tableName1 +"', 'U') is null \nBEGIN\n CREATE TABLE "+tableName1 +"\n("
                            for i in range(len(headers)):
                                headers[i]=re.sub(r"\s+", "_", re.sub(r"[^\w\s]", '',headers[i])).rstrip('_')
                                statement = (statement + '\n{} NVARCHAR({}),').format('['+headers[i]+']', 'MAX')
                            statement = statement[:-1] + '\n)\nEND'
                            cursor.execute(statement)

                        rcnt=sheet.nrows
                        ccnt=sheet.ncols
                        query = 'insert into '+dbName1+'.dbo.'+tableName1+' values ({0})'
                        query = query.format(','.join('?' * ccnt))
                        if(truncateTable1=='1'):
                            cursor.execute('TRUNCATE TABLE '+dbName1+'.dbo.'+tableName1)
                        counter=-1
                        for r in range(rcnt):
                            #print(sheet.row_values(r))
                            cursor.execute(query, sheet.row_values(r))
                            counter += 1
                        lblCounter.config(text='Inserted Record Count: '+str(counter))
                        #cursor.execute('DELETE FROM '+dbName1+'.dbo.'+tableName1+' WHERE Name='+"'Name'")
                        cursor.execute('DELETE TOP (1) FROM '+dbName1+'.dbo.'+tableName1)
                        cursor.commit()
                        messagebox.showinfo('Message', 'Data Loaded Successfully From Excel')

                    if(re.findall(r'^\d+',sheetInfo1)):
                        sheet = wb.sheet_by_index(int(sheetInfo1)-1)
                        sheet.cell_value(0, 0)
                        cursor = conn.cursor()
                        headers=[]
                        if(createTable1=='1'):
                            headers=sheet.row_values(0)
                            statement = "IF object_id('"+tableName1 +"', 'U') is null \nBEGIN\n CREATE TABLE "+tableName1 +"\n("
                            for i in range(len(headers)):
                                headers[i]=re.sub(r"\s+", "_", re.sub(r"[^\w\s]", '',headers[i])).rstrip('_')
                                statement = (statement + '\n{} NVARCHAR({}),').format('['+headers[i]+']', 'MAX')
                            statement = statement[:-1] + '\n)\nEND'
                            cursor.execute(statement)

                        rcnt=sheet.nrows
                        ccnt=sheet.ncols
                        query = 'insert into '+dbName1+'.dbo.'+tableName1+' values ({0})'
                        query = query.format(','.join('?' * ccnt))

                        if(truncateTable1=='1'):
                            cursor.execute('TRUNCATE TABLE '+dbName1+'.dbo.'+tableName1)
                        counter=-1
                        for r in range(rcnt):
                            #print(sheet.row_values(r))
                            cursor.execute(query, sheet.row_values(r))
                            counter += 1
                        lblCounter.config(text='Inserted Record Count: '+str(counter))
                        #cursor.execute('DELETE FROM '+dbName1+'.dbo.'+tableName1+' WHERE Name='+"'Name'")
                        cursor.execute('DELETE TOP (1) FROM '+dbName1+'.dbo.'+tableName1)
                        cursor.commit()
                        messagebox.showinfo('Message', 'Data Loaded Successfully From Excel')

        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'File Not Found !')
        except IOError:
            messagebox.showinfo('ERROR', 'Could not read file: ')
        except Exception as e:
            messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))
        # except ConnectionAbortedError as dbe:
        #     lblMsg.config(text='Not able to connect to DB: '+dbName1+' of the server: '+serverName1)
        #     print(dbe)
        # finally:
        #     f.close()

    if (radioBtn=='2'):
        try:

            csvfile = open(filename1, 'r',errors='ignore') #,encoding='utf8'
            if(userName1!='' and password1!=''):
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
            else:
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'File Not Found !')
        except IOError:
            messagebox.showinfo('ERROR', 'Could not read file: ')
        except Exception as e:
            messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))
        else:
            csvreader = csv.reader(csvfile)
            cur=conn.cursor() # Get the cursor
            if(truncateTable1=='1'):
                cur.execute('TRUNCATE TABLE '+dbName1+'.dbo.'+tableName1)
            idx =-1
            monthArr = []
            viewArr = []
            HeaderArr = []
            DataRow = []
            datavalue = ""
            counter=0

            Project_Name =""
            PowerSteering_ID =""
            Sequence_number	 =""
            Status	 =""
            Work_Template =""
            Work_type	=""
            Project_Lead	=""
            System_start_date	=""
            System_end_date	=""
            Currency =""
            Metric_Name =""

            for row in csvreader:
                #print(row)
                idx=idx+1
                if(idx>2):
                    DataRow = row
                    for i in range(len(DataRow)):
                        if(i<=10):
                            if(HeaderArr[i]=="Project Name"):
                                Project_Name = DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="PowerSteering ID"):
                                PowerSteering_ID=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Sequence number"):
                                Sequence_number=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Status"):
                                Status=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Work Template"):
                                Work_Template=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Work type"):
                                Work_type=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Project Lead"):
                                Project_Lead=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="System start date"):
                                System_start_date=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="System end date"):
                                System_end_date=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Currency"):
                                Currency=DataRow[i] #'"' +DataRow[i]+'"'
                            elif (HeaderArr[i]=="Metric Name"):
                                Metric_Name=DataRow[i] #'"' +DataRow[i]+'"'
                        else:

                            if(monthArr[i]=='Display Total'):
                                break
                            if(DataRow[i] ==''):
                                datavalue = "0"
                                #datavalue = "'"+"0"+'"'
                            else:
                                datavalue = DataRow[i]

                            cur.execute('INSERT INTO '+dbName1+'.dbo.'+tableName1+' VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',Project_Name,PowerSteering_ID,Sequence_number,Status,Work_Template,Work_type,Project_Lead,System_start_date,System_end_date,Currency,Metric_Name,monthArr[i],viewArr[i],HeaderArr[i],datavalue)
                            counter += 1
                            lblCounter.config(text='Inserted Record Count: '+str(counter))
                        conn.commit()

                if(idx==0):
                    monthArr = row
                    indEnd=0   #len(monthArr)
                    indLastSetSI = 0
                    #Find 1st row's 1st non blank value and insert the same value till the next non blank vlaue
                    for i in range(len(monthArr)):
                        if monthArr[i] !='':
                            indStart  =i
                            for k in range(indStart+1,len(monthArr)):
                                if monthArr[k] !='':
                                    indEnd =k
                                    break
                            for j in range(indStart+1,indEnd):
                                monthArr[j]=monthArr[i]
                            indStart=indEnd
                    #print(monthArr)
                if(idx==1):
                    viewArr = row
                    for i in range(len(viewArr)):
                        if viewArr[i] !='':
                            indStart  =i
                            for k in range(indStart+1,len(viewArr)):
                                if viewArr[k] !='':
                                    indEnd =k
                                    break
                            for j in range(indStart+1,indEnd):
                                viewArr[j]=viewArr[i]
                            indStart=indEnd
                    #print(viewArr)
                if(idx==2):
                    HeaderArr = row

            messagebox.showinfo('Message', 'TDC Data Loaded Successfully')

    if (radioBtn=='3'):
        try:
            if not os.path.isdir(filename1):
                raise NotADirectoryError
            #file1 = open("Attachment.csv","w",encoding='utf8')
            if(userName1!='' and password1!=''):
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
            else:
                conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

            cursor = conn.cursor()
            if(truncateTable1=='1'):
                cursor.execute('TRUNCATE TABLE '+dbName1+'.dbo.'+tableName1)


            folder_size = 0
            for (path, dirs, files) in os.walk(filename1):
                for file in files:
                    filenamewithloc = os.path.join(path, file)
                    folder_size = os.path.getsize(filenamewithloc)

                    #print(file+' '+str(round(folder_size/(1024.0))))
                    dataStr = file+'^'+str(round(folder_size/(1024.0)))#+'\n'
                    #file1.write(dataStr)
                    dataStr = dataStr.split('^')
                    cursor.execute("INSERT INTO "+dbName1+".dbo."+tableName1+" VALUES (?,?,?,?,?)",dataStr)
            conn.commit()
            messagebox.showinfo('Message', 'Attatchment Data Loaded Successfully')

        except FileNotFoundError:
            messagebox.showinfo('ERROR', 'File Not Found !')
        except IOError:
            messagebox.showinfo('ERROR', 'Could not read file: ')
        except Exception as e:
            messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))


bt = tk.Button(window, text="Load Data \rRun Script",bg="red", fg="yellow",width = 10, height = 2, bd = 5,command=clicked, cursor="hand2")
bt.grid (column=2, row=27)

bt1 = tk.Button(window, text="...", command=load_file, width=2,height = 1)
bt1.grid(column=4,row=6)

helplink1 = tk.Label(window, text="Help", fg="blue", cursor="hand2")
helplink1.grid (column=4, row=0)
helplink1.bind("<Button-1>", lambda e: callback("http://www.google.com"))


window.geometry('560x400')
window.mainloop()