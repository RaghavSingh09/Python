import tkinter as tk                # python 3
from tkinter import font  as tkfont # python 3
from tkinter import messagebox,filedialog
from tkinter.ttk import *
from PIL import Image, ImageTk
import pyodbc as pdb
import csv
import os
import xlrd
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror
import re
import cx_Oracle
import datetime

class DataTransferApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")

        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (StartPage, PageSql, PageOracle,PageMySql,PageSqlOracle,PageSqlMySql):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()


class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="Welcome To Data Loader Tool", font=controller.title_font,fg="blue")
        label.pack(side="top", fill="x", pady=10)

        button1 = tk.Button(self, text="SQL SERVER",font=("Arial Bold",10),bg="red", fg="white",
                            command=lambda: controller.show_frame("PageSql"))
        button2 = tk.Button(self, text="ORACLE",font=("Arial Bold",10),bg="white", fg="red",
                            command=lambda: controller.show_frame("PageOracle"))
        button3 = tk.Button(self, text="MY-SQL",font=("Arial Bold",10),bg="red", fg="white",
                            command=lambda: controller.show_frame("PageMySql"))
        button4 = tk.Button(self, text="SQL SERVER To/From ORACLE",font=("Arial Bold",10),bg="white", fg="red",
                            command=lambda: controller.show_frame("PageSqlOracle"))
        button5 = tk.Button(self, text="SQL SERVER To/From MY-SQL",font=("Arial Bold",10),bg="red", fg="white",
                            command=lambda: controller.show_frame("PageSqlMySql"))
        button1.pack(side = "top", expand = True, fill = "both")
        button2.pack(side = "top", expand = True, fill = "both")
        button3.pack(side = "top", expand = True, fill = "both")
        button4.pack(side = "top", expand = True, fill = "both")
        button5.pack(side = "top", expand = True, fill = "both")

        label = tk.Label(self, text="\n", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)


class PageSql(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        #label = tk.Label(self, text="Working In SQL Server Data Transfer Mode ", font=("Arial Bold", 10))
        #label.pack(side="top", fill="x", pady=10)
        #label.grid(column=0,row=2)
        lblFileName = tk.Label(self, text="Enter Filename With Folder Location:", font=("Arial", 10))
        lblServerAdd = tk.Label(self, text="Enter SQL Server Address:             ", font=("Arial", 10))
        lblDBName = tk.Label(self, text="Enter Database Name To Connect:  ", font=("Arial", 10))
        lblUN = tk.Label(self, text="Enter SQL Server Username:          ", font=("Arial", 10))
        lblPWD = tk.Label(self, text="Enter SQL Server Password:           ", font=("Arial", 10))
        lblCommand = tk.Label(self, text="Enter Table Name \ SQL Command:", font=("Arial", 10))
        lblMsg = tk.Label(self, text="", font=("Arial Bold", 10),fg='blue')
        lblSNCV = tk.Label(self, text="SQL Native Client Version: ", font=("Arial", 10))
        lblCounter = tk.Label(self,text='',font=("Arial", 10),fg="green")
        lblExcelInfo = tk.Label(self, text="Enter Excel Sheet Number \ Name:  ", font=("Arial", 10))
        lblAction = tk.Label(self, text="\nActions To Perform  ", font=("Arial Bold", 8))
        lblLoadType = tk.Label(self, text="\nData Transfer Mode", font=("Arial Bold", 8))
        
        lblFileName.grid (column=0, row=6)
        lblExcelInfo.grid (column=0, row=8)
        lblServerAdd.grid (column=0, row=10)
        lblDBName.grid (column=0, row=12)
        lblUN.grid (column=0, row=14)
        lblPWD.grid (column=0, row=16)
        lblCommand.grid (column=0, row=18)
        lblSNCV.grid (column=0, row=20)
        lblAction.grid (column=0, row=22)
        lblLoadType.grid (column=3, row=22)
        lblMsg.grid (column=0, row=24)
        lblCounter.grid(column=3,row=28)

        entryText = tk.StringVar()
        varID = tk.IntVar()
        varRS = tk.IntVar()
        varCT = tk.IntVar()
        varTT = tk.IntVar()
        varTT.set(1)
        var = tk.IntVar()
        
        filename = tk.Entry(self,width=30,textvariable=entryText )
        filename.grid(column=3, row=6)
        sheetinfo = tk.Entry(self,width=30)
        sheetinfo.grid(column=3, row=8)
        serverName = tk.Entry(self,width=30)
        serverName.grid(column=3, row=10)
        dbName = tk.Entry(self,width=30)
        dbName.grid(column=3, row=12)
        userName = tk.Entry(self,width=30)
        userName.grid(column=3, row=14)
        password = tk.Entry(self,width=30)
        password.grid(column=3, row=16)
        tableName = tk.Text(self, font=("Arial", 9),width=26,height=4)
        tableName.grid(column=3, row=18)
        nativeCV = tk.Spinbox(self, from_ = 11, to = 20,width=10)
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
            fname = askopenfilename(filetypes=(("All files", "*.*"),("Excel files", "*.xlsx"),("CSV files", "*.csv"),("SQL files", "*.sql"),("Text files", "*.txt")))
            if fname:
                try:
                    entryText.set(fname)
                except:
                    showerror("Open Source File", "Failed to read file\n'%s'" % fname)
                return

        def callback(url):
            #webbrowser.open_new(url)
            #webbrowser.open_new(r"file://c:\test\test.csv")
            help='''***Check Boxes***\nID - Import Data To CSV\nTT - Truncate Table\nCT - Create Table\nRS - Run Script\n\n***Radio Buttons***
                    Non-TDC - Normal CSV or Excel File\nTDC - GoodYear Template TDC File in CSV Format\nAttachment - Load attachment files from a folder\n... - Browse Files To Load Data\n
***Conditions***\n
1. If loading data from excel file then mention workbook/sheet number/name.\n
2. If you want run a SQL command then enter it in SQL command text box and select RS check box.\n
3. If you are using selfs authenticaton then SQL UN and PWD is not required.\n
***Exceptions***\n
1. CSV files will not be able to load multi-language data, use excel file.\n
2. If excel file contains date column then please update those columns in your table to get correct date.
   DATEADD(Day,CAST(Your_Col_Name AS INT)-2,'1900-01-01')\n
3. If you are running a SQL script file then make sure it contains Go after each SQL block and select RS check box.\n'''
            messagebox.showinfo('Help Message',help )

        nonTdcType = tk.Radiobutton(self, text='Import Data To DB  ', variable=var, value=1, command=sel)
        importData = tk.Radiobutton(self, text='Export Data To File  ', variable=var, value=2, command=sel)
        loadAttach = tk.Radiobutton(self, text='Load Attachment    ', variable=var, value=3, command=sel)
        truncateTable = tk.Checkbutton(self, text='Truncate Table ', variable=varTT)
        createTable = tk.Checkbutton(self, text='Create Table     ', variable=varCT)
        runScript = tk.Checkbutton(self, text='Run Script         ', variable=varRS)
        #importData = tk.Checkbutton(self, text='ID', variable=varID)
        
        nonTdcType.grid(column=3, row=23)
        importData.grid(column=3, row=24)
        loadAttach.grid(column=3, row=25)
        truncateTable.grid(column=0, row=23)
        createTable.grid(column=0, row=24)
        runScript.grid(column=0, row=25)

        
        def get_spinbox_value():
            sbv=str(nativeCV.get())
            return sbv
        
        def runDataLoaderApp():
            filename1 = filename.get()
            sheetInfo1=sheetinfo.get()
            serverName1 = serverName.get()
            dbName1 = dbName.get()
            userName1 = userName.get()
            password1 = password.get()
            tableName1 = tableName.get("1.0",'end-1c')
            radioBtn = sel()
            spinBox = get_spinbox_value()
            truncateTable1=varTT_states()
            createTable1=varCT_states()
            runScript1=varRS_states()

            if(tableName1.find('.') != -1):
                tableName1=tableName1
            else:
                tableName1 = 'dbo.'+tableName1

            if (radioBtn=='0' and runScript1=='0'):
                #lblMsg.config(text='Error Msg:\n' +'Please Select File Load Type !',fg="red")
                messagebox.showinfo('ERROR', 'Please Select File Load Type !')

            if (radioBtn=='2' and filename1!=''):
                try:

                    if(userName1!='' and password1!=''):
                        conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
                    else:
                        conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

                    cursor = conn.cursor()
                    cursor.execute(tableName1)
                    if(filename1.endswith(".csv") or filename1.endswith(".txt")):
                        with open(filename1, "w", newline='') as f:
                            wrtr = csv.writer(f)
                            wrtr.writerow([i[0] for i in cursor.description]) # write headers
                            for row in cursor:
                                wrtr.writerow(row)

                        messagebox.showinfo('Message', 'Data Exported Successfully')
                    else:
                        messagebox.showinfo('Error', 'Please Select CSV/TXT File To Import Data')


                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))

            if (runScript1=='1' and filename1==''):
                try:

                    if(userName1!='' and password1!=''):
                        conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'uid='+userName1+';pwd='+password1+'')
                    else:
                        conn = pdb.connect('Driver={SQL Server Native Client '+spinBox+'.0};' 'Server='+serverName1+';' 'Database='+dbName1+';' 'Trusted_Connection=yes')

                    cursor = conn.cursor()
                    cursor.execute(tableName1)
                    cursor.commit()

                    messagebox.showinfo('Message', 'Query Executed Successfully')


                    if(radioBtn=='2'):
                        cursor.execute(tableName1)
                        if(filename1.endswith(".csv") or filename1.endswith(".txt")):
                            with open(filename1, "w", newline='') as f:
                                wrtr = csv.writer(f)
                                for row in cursor:
                                    wrtr.writerow(row)

                            messagebox.showinfo('Message', 'Data Exported Successfully')
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

                        query = 'insert into '+dbName1+'.'+tableName1+' values ({0})'
                        query = query.format(','.join('?' * len(data)))

                        if(truncateTable1=='1'):
                            cursor.execute('TRUNCATE TABLE '+dbName1+'.'+tableName1)
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
                            query = 'insert into '+dbName1+'.'+tableName1+' values ({0})'
                            query = query.format(','.join('?' * ccnt))
                            if(truncateTable1=='1'):
                                cursor.execute('TRUNCATE TABLE '+dbName1+'.'+tableName1)
                            counter=-1
                            for r in range(rcnt):
                                #print(sheet.row_values(r))
                                cursor.execute(query, sheet.row_values(r))
                                counter += 1
                            lblCounter.config(text='Inserted Record Count: '+str(counter))
                            #cursor.execute('DELETE FROM '+dbName1+'.'+tableName1+' WHERE Name='+"'Name'")
                            cursor.execute('DELETE TOP (1) FROM '+dbName1+'.'+tableName1)
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
                            query = 'insert into '+dbName1+'.'+tableName1+' values ({0})'
                            query = query.format(','.join('?' * ccnt))

                            if(truncateTable1=='1'):
                                cursor.execute('TRUNCATE TABLE '+dbName1+'.'+tableName1)
                            counter=-1
                            for r in range(rcnt):
                                #print(sheet.row_values(r))
                                cursor.execute(query, sheet.row_values(r))
                                counter += 1
                            lblCounter.config(text='Inserted Record Count: '+str(counter))
                            #cursor.execute('DELETE FROM '+dbName1+'.'+tableName1+' WHERE Name='+"'Name'")
                            cursor.execute('DELETE TOP (1) FROM '+dbName1+'.'+tableName1)
                            cursor.commit()
                            messagebox.showinfo('Message', 'Data Loaded Successfully From Excel')

                except FileNotFoundError:
                    messagebox.showinfo('ERROR', 'File Not Found !')
                except IOError:
                    messagebox.showinfo('ERROR', 'Could not read file: ')
                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))

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
                        cursor.execute('TRUNCATE TABLE '+dbName1+'.'+tableName1)


                    folder_size = 0
                    for (path, dirs, files) in os.walk(filename1):
                        for file in files:
                            filenamewithloc = os.path.join(path, file)
                            folder_size = os.path.getsize(filenamewithloc)

                            #print(file+' '+str(round(folder_size/(1024.0))))
                            dataStr = file+'^'+str(round(folder_size/(1024.0)))#+'\n'
                            #file1.write(dataStr)
                            #dataStr = dataStr.split('^')
                            cursor.execute("INSERT INTO "+dbName1+"."+tableName1+" VALUES (?)",dataStr)
                    conn.commit()
                    messagebox.showinfo('Message', 'Attatchment Data Loaded Successfully')

                except FileNotFoundError:
                    messagebox.showinfo('ERROR', 'File Not Found !')
                except IOError:
                    messagebox.showinfo('ERROR', 'Could not read file: ')
                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))

        bt1 = tk.Button(self, text="...", command=load_file, width=2,height = 1)
        bt1.grid(column=4,row=6)
        btnLoad = tk.Button(self, text="LOAD DATA",bg="red", fg="white",command=runDataLoaderApp)
        btnHome = tk.Button(self, text="<--HOME",bg="black", fg="white",command=lambda: controller.show_frame("StartPage"))
        helplink1 = tk.Label(self, text="Help", fg="blue", cursor="hand2")
        helplink1.grid (column=4, row=0)
        helplink1.bind("<Button-1>", lambda e: callback("http://www.google.com"))
        btnHome.grid(column=0,row=27)
        btnLoad.grid(column=2,row=27)


class PageOracle(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        entryText = tk.StringVar()
        varED = tk.IntVar()
        varID = tk.IntVar()
        varRS = tk.IntVar()
        varCT = tk.IntVar()
        varTT = tk.IntVar()
        varTT.set(1)
        var = tk.IntVar()
        banner = tk.Label(self, text="Working In Oracle Data Transfer Mode    \n", font=("Arial Bold", 10))
        fileInfo = tk.Label(self, text="Enter Filename With Folder Location: ", font=("Arial", 10))
        loginInfo = tk.Label(self, text="Enter Oracle Login Information:          ", font=("Arial", 10))
        lblLoginInfo = tk.Label(self, text="Format: UN/PWD@IP:Port/SID \n UN/PWD@IP/ServiceName", font=("Arial Bold", 8))
        excelInfo = tk.Label(self, text="Enter Excel Sheet Number \ Name:    ", font=("Arial", 10))
        commandInfo=tk.Label(self, text="Enter Table Name \ SQL Command: ", font=("Arial", 10))
        lblCounter = tk.Label(self,text='',font=("Arial", 10),fg="green")
        filename = tk.Entry(self,width=30,textvariable=entryText )
        excelSheetInfo = tk.Entry(self,width=30)
        serverInfo = tk.Entry(self,width=30)
        commandText = tk.Text(self, font=("Arial", 9),width=26,height=4)
        lblAction = tk.Label(self, text="\nActions To Perform  ", font=("Arial Bold", 8))
        lblLoadType = tk.Label(self, text="\nData Transfer Mode", font=("Arial Bold", 8))

        createTable = tk.Checkbutton(self, text='Create Table    ', variable=varCT)
        truncateTable = tk.Checkbutton(self, text='Truncate Table', variable=varTT)
        runScript = tk.Checkbutton(self, text='Run Script        ', variable=varRS)

        def varCT_states():
            ct=str(varCT.get())
            return ct

        def varTT_states():
            tt=str(varTT.get())
            return tt

        def varRS_states():
            rs=str(varRS.get())
            return rs

        def sel():
            rop=str(var.get())
            return rop

        def load_file():
            fname = askopenfilename(filetypes=(("All files", "*.*"),("Excel files", "*.xlsx"),("CSV files", "*.csv"),("SQL files", "*.sql"),("Text files", "*.txt")))
            if fname:
                try:
                    entryText.set(fname)
                    #filename.config(textvariable=entryText,fg="red")
                except:
                    showerror("Open Source File", "Failed to read file\n'%s'" % fname)
                return

        exportData = tk.Radiobutton(self, text='Export DB Data To Files ', variable=var, value=2, command=sel)
        importData = tk.Radiobutton(self, text='Import Data To Oracle  ', variable=var, value=1, command=sel)

        def callback(url):
            #webbrowser.open_new(url)
            #webbrowser.open_new(r"file://c:\test\test.csv")
            help='''***Oracle Data Transfer Help Log***\n'''
            messagebox.showinfo('Help Message',help )

        def runDataLoaderApp():
            filename1 = filename.get()
            excelSheetInfo1=excelSheetInfo.get()
            serverInfo1=serverInfo.get()
            commandText1 = commandText.get("1.0",'end-1c')#commandText.get()
            truncateTable1=varTT_states()
            createTable1=varCT_states()
            runScript1=varRS_states()
            importData1=sel()

            if (runScript1=='0' and importData1=='0'):
                messagebox.showinfo('ERROR', 'Please Select File Load Type !')
            if (importData1=='2' and filename1!=''):#Import Data Logic
                try:
                    con = cx_Oracle.connect(serverInfo1)#MYHR/MYHR@127.0.0.1:1521/XE
                    cursor = con.cursor()
                    if(filename1.endswith(".csv") or filename1.endswith(".txt")):
                        csv_file = open(filename1, "w")
                        writer = csv.writer(csv_file, delimiter=',', lineterminator="\n", quoting=csv.QUOTE_NONNUMERIC)
                        cursor.execute(commandText1)
                        writer.writerow([i[0] for i in cursor.description])
                        for row in cursor:
                            writer.writerow(row)

                        messagebox.showinfo('Message', 'Data Exported Successfully.')
                    else:
                        messagebox.showinfo('Error', 'Please Select Csv/Text File To Export Data.')

                    cursor.close()
                    con.close()
                    csv_file.close()
                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))

            if(importData1=='1'):
                try:
                    conn = cx_Oracle.connect(serverInfo1)#MYHR/MYHR@127.0.0.1:1521/XE

                    if(filename1.endswith("csv")):
                        f=open (filename1, 'r',errors='ignore') #,encoding='utf8',errors='ignore'
                        reader = csv.reader(f)
                        headers=[]
                        L = []
                        column_list=''
                        value_list=''
                        first_row = next(reader)
                        cursor = conn.cursor()

                        if(createTable1=='1'):
                            headers=first_row
                            statement = "CREATE TABLE "+commandText1 +"\n("
                            for i in range(len(headers)):
                                headers[i]=re.sub(r"\s+", "_", re.sub(r"[^\w\s]", '',headers[i])).rstrip('_')
                                statement = (statement + '\n{} NVARCHAR2({}),').format('"'+headers[i]+'"', '1000')
                            statement = statement[:-1] + '\n)'
                            cursor.execute(statement)

                        if(truncateTable1=='1'):
                            cursor.execute('TRUNCATE TABLE '+commandText1)

                        column_string = ','.join(first_row)
                        insert_string='insert into ' + commandText1 + ' (' + column_string + ') values ('
                        val_list=[]
                        for i in range(1,len(first_row)+1):
                            val_list.append(':'+ str(i))
                        value_string=','.join(val_list)
                        insert_string += value_string + ')'
                        for row in reader:
                            for index,col in enumerate(row):
                                col_tr = col
                                if col_tr:
                                    if col_tr[0] != '"' :
                                        try:
                                            col_tr=datetime.datetime.strptime(col_tr,'%d-%b-%y')
                                        except ValueError:
                                            continue
                                row[index] = col_tr
                            L.append(row)
                        cursor.prepare(insert_string)
                        cursor.executemany(None, L)
                        lblCounter.config(text='Inserted Record Count: '+str(cursor.rowcount))
                        #print('Inserted: ' + str(cursor.rowcount) + ' rows.')
                        conn.commit()
                        #cursor.commit()
                        messagebox.showinfo('Message', 'Data Loaded Successfully From CSV')

                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))



        btnLoad = tk.Button(self, text="Load Data",bg="red", fg="white",command=runDataLoaderApp)
        btnHome = tk.Button(self, text="<--HOME",bg="black", fg="white",command=lambda: controller.show_frame("StartPage"))
        btnLoadFile = tk.Button(self, text="...", command=load_file, width=2,height = 1)
        helplink1 = tk.Label(self, text="Help", fg="blue", cursor="hand2")
        helplink1.grid (column=2, row=2)
        helplink1.bind("<Button-1>", lambda e: callback("http://www.google.com"))

        banner.grid(column=0,row=2)
        fileInfo.grid(column=0, row=4)
        filename.grid(column=1, row=4)
        excelInfo.grid(column=0, row=5)
        excelSheetInfo.grid(column=1, row=5)
        loginInfo.grid(column=0,row=6)
        serverInfo.grid(column=1, row=6)
        lblLoginInfo.grid(column=1, row=7)
        commandInfo.grid(column=0,row=8)
        commandText.grid(column=1,row=8)
        lblAction.grid(column=0,row=9)
        lblLoadType.grid(column=1,row=9)
        truncateTable.grid(column=0, row=10)
        createTable.grid(column=0, row=12)
        exportData.grid(column=1, row=12)
        importData.grid(column=1, row=10)
        runScript.grid(column=0, row=14)
        lblCounter.grid(column=1,row=25)
        btnLoadFile.grid(column=2,row=4)
        btnHome.grid(column=0,row=20)
        btnLoad.grid(column=1,row=20)

class PageMySql(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="                       Transfer Data Between CSV, Flat-Files, Excel and MY-SQL DB", font=("Arial Bold", 10))
        #label.pack(side="top", fill="x", pady=10)
        label.grid(column=2,row=2)
        button = tk.Button(self, text="<--HOME",bg="black", fg="white",command=lambda: controller.show_frame("StartPage"))
        button.grid(column=2,row=5)

class PageSqlOracle(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        entryText = tk.StringVar()
        varRS = tk.IntVar()
        varCT = tk.IntVar()
        varTT = tk.IntVar()
        #varTT.set(1)
        var = tk.IntVar()
        #SQL Server Information
        lblSqlServerInfo = tk.Label(self, text="SQL Server Information:             ", font=("Arial Bold", 12))
        lblSqlServerAdd = tk.Label(self, text="Enter SQL Server Address:             ", font=("Arial", 10))
        lblDBName = tk.Label(self, text="Enter SQL Database Name:            ", font=("Arial", 10))
        lblUN = tk.Label(self, text="Enter SQL Server Username:           ", font=("Arial", 10))
        lblPWD = tk.Label(self, text="Enter SQL Server Password:           ", font=("Arial", 10))
        lblSqlCommand = tk.Label(self, text="Enter Table Name \ SQL Command:", font=("Arial", 10))
        lblMsg = tk.Label(self, text="", font=("Arial Bold", 10),fg='blue')
        lblSNCV = tk.Label(self, text="SQL Native Client Version: ", font=("Arial", 10))

        lblSqlServerInfo.grid (column=0, row=8)
        lblSqlServerAdd.grid (column=0, row=10)
        lblDBName.grid (column=0, row=12)
        lblUN.grid (column=0, row=14)
        lblPWD.grid (column=0, row=16)
        lblSqlCommand.grid (column=0, row=18)
        lblSNCV.grid (column=0, row=20)

        sqlServerName = tk.Entry(self,width=30)
        sqlServerName.grid(column=1, row=10)
        sqldbName = tk.Entry(self,width=30)
        sqldbName.grid(column=1, row=12)
        sqlUserName = tk.Entry(self,width=30)
        sqlUserName.grid(column=1, row=14)
        sqlPassword = tk.Entry(self,width=30)
        sqlPassword.grid(column=1, row=16)
        sqlCommand = tk.Text(self, font=("Arial", 9),width=26,height=3)
        sqlCommand.grid(column=1, row=18)
        sqlnativeCV = tk.Spinbox(self, from_ = 11, to = 20,width=10)
        sqlnativeCV.grid(column=1,row=20)

        #Oracle Information
        lblOracleServerInfo = tk.Label(self, text="Oracle Server Information:        ", font=("Arial Bold", 12))
        lblOracleLoginInfo = tk.Label(self, text="Enter Oracle Login Information:          ", font=("Arial", 10))
        lblLoginInfoFormat = tk.Label(self, text="Format: UN/PWD@IP:Port/SID \n UN/PWD@IP/ServiceName", font=("Arial Bold", 8))
        lblOracleCommand=tk.Label(self, text="Enter Table Name \ SQL Command: ", font=("Arial", 10))
        oracleCommand = tk.Text(self, font=("Arial", 9),width=26,height=3)
        oracleServerInfo = tk.Entry(self,width=30)
        lblOracleServerInfo.grid (column=0, row=22)
        lblOracleLoginInfo.grid(column=0,row=24)
        oracleServerInfo.grid(column=1, row=24)
        lblLoginInfoFormat.grid(column=1, row=25)
        lblOracleCommand.grid(column=0,row=26)
        oracleCommand.grid(column=1,row=26)

        #Actions and Load Type
        lblAction = tk.Label(self, text="\nActions To Perform  ", font=("Arial Bold", 8))
        lblLoadType = tk.Label(self, text="\nData Transfer Mode", font=("Arial Bold", 8))

        createTable = tk.Checkbutton(self, text='Create Table    ', variable=varCT)
        truncateTable = tk.Checkbutton(self, text='Truncate Table', variable=varTT)
        #runScript = tk.Checkbutton(self, text='Run Script        ', variable=varRS)
        lblAction.grid(column=0,row=27)
        lblLoadType.grid(column=1,row=27)
        truncateTable.grid(column=0, row=28)
        createTable.grid(column=0, row=29)
        #runScript.grid(column=0, row=30)
        lblCounter = tk.Label(self,text='',font=("Arial", 10),fg="green")
        lblCounter.grid(column=1,row=36)

        def get_spinbox_value():
            sbv=str(sqlnativeCV.get())
            return sbv

        def varCT_states():
            ct=str(varCT.get())
            return ct

        def varTT_states():
            tt=str(varTT.get())
            return tt

        def varRS_states():
            rs=str(varRS.get())
            return rs

        def sel():
            rop=str(var.get())
            return rop

        sqlToOracle = tk.Radiobutton(self, text='SQL Server To Oracle', variable=var, value=1, command=sel)
        oracleToSql = tk.Radiobutton(self, text='Oracle To SQL Server', variable=var, value=2, command=sel)
        sqlToOracle.grid(column=1, row=28)
        oracleToSql.grid(column=1, row=29)

        #Data Transfer Logic

        def runDataLoaderApp():

            #Get All The Inputs From The Screen
            sqlServerName1=sqlServerName.get()
            sqldbName1 = sqldbName.get()
            sqlUserName1 = sqlUserName.get()
            sqlPassword1 = sqlPassword.get()
            sqlNCV = get_spinbox_value()
            sqlCommand1 = sqlCommand.get("1.0",'end-1c')#sqlCommand.get()
            if(tableName1.find('.') != -1):
                sqlCommand1=sqlCommand1
            else:
                sqlCommand1 = 'dbo.'+sqlCommand1

            oracleServerInfo1=oracleServerInfo.get()
            oracleCommand1 = oracleCommand.get("1.0",'end-1c')#oracleCommand.get()
            truncateTable1=varTT_states()
            createTable1=varCT_states()
            #runScript1=varRS_states()
            dataTransMode=sel()

            if(dataTransMode=='0'):
                messagebox.showinfo('ERROR', 'Please Select File Load Type !')
            if(dataTransMode=='1'):#SQL TO Oracle
                try:
                    if(sqlUserName1!='' and sqlPassword1!=''):
                        sqlConn = pdb.connect('Driver={SQL Server Native Client '+sqlNCV+'.0};' 'Server='+sqlServerName1+';' 'Database='+sqldbName1+';' 'uid='+sqlUserName1+';pwd='+sqlPassword1+'')
                    else:
                        sqlConn = pdb.connect('Driver={SQL Server Native Client '+sqlNCV+'.0};' 'Server='+sqlServerName1+';' 'Database='+sqldbName1+';' 'Trusted_Connection=yes')

                    oracleConn = cx_Oracle.connect(oracleServerInfo1)#MYHR/MYHR@127.0.0.1:1521/XE

                    sqlCursor = sqlConn.cursor()
                    oracleCursor = oracleConn.cursor()

                    sqlCursor.execute(sqlCommand1)
                    first_row = ([i[0] for i in sqlCursor.description])

                    if(createTable1=='1'):
                        headers=first_row
                        statement = "CREATE TABLE "+oracleCommand1 +"\n("
                        for i in range(len(headers)):
                            headers[i]=re.sub(r"\s+", "_", re.sub(r"[^\w\s]", '',headers[i])).rstrip('_')
                            statement = (statement + '\n{} NVARCHAR2({}),').format('"'+headers[i]+'"', '1000')
                        statement = statement[:-1] + '\n)'
                        oracleCursor.execute(statement)

                    if(truncateTable1=='1'):
                        oracleCursor.execute('TRUNCATE TABLE '+oracleCommand1)

                    query = 'insert into '+oracleCommand1+' values ({0})'
                    query = query.format(','.join('?' * len(first_row)))
                    print(query)
                    counter=0
                    for data in sqlCursor:
                        print(data)
                        oracleCursor.execute(query, data)
                        counter += 1
                        lblCounter.config(text='Inserted Record Count: '+str(counter))

                    #oracleConn.commit()
                    #oracleConn.close()
                    '''
                    column_string = ','.join(first_row)
                    insert_string='INSERT INTO ' + oracleCommand1 + ' (' + column_string + ') VALUES ('
                    L = []
                    val_list=[]
                    counter=0
                    for i in range(1,len(first_row)+1):
                        val_list.append(':'+ str(i))
                    value_string=','.join(val_list)
                    insert_string += value_string + ')'

                    for row in sqlCursor:
                        counter += 1
                        for index,col in enumerate(row):
                            col_tr = col
                            if col_tr:
                                if col_tr[0] != '"' :
                                    try:
                                        col_tr=datetime.datetime.strptime(col_tr,'%d-%b-%y')
                                    except ValueError:
                                        continue
                            row[index] = col_tr
                            print(row)
                        L.append(row)
                    print('Outer Yes')
                    oracleCursor.prepare(insert_string)
                    #oracleCursor.executemany(None, L)
                    print(insert_string)
                    oracleConn.commit()
                    oracleConn.close()
                    lblCounter.config(text='Inserted Record Count: '+str(counter))
                    '''
                    messagebox.showinfo('Message', 'Data Loaded Successfully From Sql Server To Oracle DB !')
                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))

            if(dataTransMode=='2'):
                try:
                    if(sqlUserName1!='' and sqlPassword1!=''):
                        sqlConn = pdb.connect('Driver={SQL Server Native Client '+sqlNCV+'.0};' 'Server='+sqlServerName1+';' 'Database='+sqldbName1+';' 'uid='+sqlUserName1+';pwd='+sqlPassword1+'')
                    else:
                        sqlConn = pdb.connect('Driver={SQL Server Native Client '+sqlNCV+'.0};' 'Server='+sqlServerName1+';' 'Database='+sqldbName1+';' 'Trusted_Connection=yes')

                    oracleConn = cx_Oracle.connect(oracleServerInfo1)#MYHR/MYHR@127.0.0.1:1521/XE

                    sqlCursor = sqlConn.cursor()
                    oracleCursor = oracleConn.cursor()

                    oracleCursor.execute(oracleCommand1)

                    data = ([i[0] for i in oracleCursor.description])

                    if(createTable1=='1'):
                        headers=data
                        statement = "IF object_id('"+sqlCommand1 +"', 'U') is null \nBEGIN\n CREATE TABLE "+sqlCommand1 +"\n("
                        for i in range(len(headers)):
                            headers[i]=re.sub(r"\s+", "_", re.sub(r"[^\w\s]", '',headers[i])).rstrip('_')
                            statement = (statement + '\n{} NVARCHAR({}),').format('['+headers[i]+']', 'MAX')
                        statement = statement[:-1] + '\n)\nEND'
                        sqlCursor.execute(statement)

                    query = 'insert into '+sqldbName1+'.'+sqlCommand1+' values ({0})'
                    query = query.format(','.join('?' * len(data)))

                    if(truncateTable1=='1'):
                        sqlCursor.execute('TRUNCATE TABLE '+sqldbName1+'.'+sqlCommand1)

                    counter=0
                    for data in oracleCursor:
                        sqlCursor.execute(query, data)
                        counter += 1
                        lblCounter.config(text='Inserted Record Count: '+str(counter))

                    sqlCursor.commit()
                    messagebox.showinfo('Message', 'Data Loaded Successfully From Oracle To Sql Server DB !')
                except Exception as e:
                    messagebox.showinfo('ERROR', 'Someting Went Wrong ! \n'+str(e))


        btnLoad = tk.Button(self, text="Load Data",bg="red", fg="white",command=runDataLoaderApp)
        btnHome = tk.Button(self, text="<--HOME",bg="black", fg="white",command=lambda: controller.show_frame("StartPage"))
        btnLoad.grid(column=1,row=35)
        btnHome.grid(column=0,row=35)


class PageSqlMySql(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="                       Transfer Data Between Sql Server and MY-SQL DB", font=("Arial Bold", 10))
        #label.pack(side="top", fill="x", pady=10)
        label.grid(column=2,row=2)
        button = tk.Button(self, text="<--HOME",bg="black", fg="white",command=lambda: controller.show_frame("StartPage"))
        button.grid(column=2,row=5)

if __name__ == "__main__":
    app = DataTransferApp()
    app.title("Python Data Loader Application")
    app.resizable(0, 0)
    #app.geometry('560x400')
    app.geometry('520x480')
    app.mainloop()


