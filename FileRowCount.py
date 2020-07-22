import csv
import os
import xlrd
import xlwt
from xlwt import Workbook

sheetName = "0"
inputFile = r"C:\Users\rkumar699\Desktop\info.xlsx"
outputFile = r"C:\Users\rkumar699\Desktop\output.xls"

def fn_FileRowCount(fileName):
    if fileName.find(".xls")==-1:
        with open(fileName,"r") as fr:
            reader = csv.reader(fr,delimiter = ",")
            data = list(reader)
            row_count = len(data)
            #print(fileName+" Row Count is: "+str(row_count))
            return fileName,str(row_count)
    else:
        wb = xlrd.open_workbook(fileName)
        try:
            sheet = wb.sheet_by_name(sheetName)
        except:
            sheet = wb.sheet_by_index(int(sheetName))
        row_count=sheet.nrows
        #print(fileName+" Row Count is: "+str(row_count))
        return fileName,str(row_count)

#Reading Folder Loactions From Excel File
inputWB = xlrd.open_workbook(inputFile)
excelSheet = inputWB.sheet_by_index(int(sheetName))
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
      fn, rc =  fn_FileRowCount(fileLoc)
      outputSheet.write(outputRc,0,fn)
      outputSheet.write(outputRc,1,rc)
      outputRc += 1
    elif os.path.isdir(fileLoc):
        for root, directories, files in os.walk(fileLoc, topdown=False):
            for name in files:
                fn, rc =  fn_FileRowCount(os.path.join(root, name))
                outputSheet.write(outputRc,0,fn)
                outputSheet.write(outputRc,1,rc)
                outputRc += 1
    else:
        print("File or Directory Does Not Exists")

outputWB.save(outputFile)

