import pandas as pd
import pypyodbc as pdb
import win32com.client as mycom
from win32com.client import constants as c
import os
import shutil
import sys
from datetime import datetime, timedelta
# import csv
from pathlib import Path
# from dateutil.relativedelta import relativedelta


# ============== VARIABLES
workdir = "E:\\03_REPORT\\01_DAILY\\Inventory_FootLocker"
# workdir = "D:\\Ismail_MAA\\apps"
dt_aod = datetime.strptime(str(sys.argv[1]), '%Y%m%d') #As of date
dt_ed = dt_aod - timedelta(days=1)
dt_raw = datetime.strftime(dt_ed, '%Y%m%d')
dt_file = datetime.strftime(dt_aod, '%Y%m%d')

#============ CREATE EXCEL
# bawah  buat abis restart server
xl = mycom.Dispatch('Excel.Application')
# xl = mycom.gencache.EnsureDispatch('Excel.Application')

resultExcel = os.path.join(workdir,'RawData_Inv_Footlocker_' + dt_raw.replace(' ', '') +'.xlsb')
# 2022-03-31 --> modif penamaan file
# resultCsv = os.path.join(workdir,'FL' + dt_file.replace(' ', '') +'.INV.csv')
# resultSls = os.path.join(workdir,'FL' + dt_file.replace(' ', '') +'.INV.sls')
resultCsv = os.path.join(workdir,'FLINV' + dt_file.replace(' ', '') +'.csv')
resultSls = os.path.join(workdir,'FLINV' + dt_file.replace(' ', '') +'.sls')
# print(resultExcel)
xl.Visible = False

wb = xl.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
ws.Name = 'WPInv'
xl.ActiveWindow.DisplayGridlines = False
xl.DisplayAlerts = False

cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
q = "EXEC USP_FOOTLOCKER 9,'" + str(sys.argv[1]) + "'"
df = pd.read_sql(q, cn)

# Paste header
st_row = 1
st_col = 1
for col in df.columns:
  ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
  st_col = st_col + 1

# Paste recordset
st_row = 2
st_col = 1
ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
         ws.Cells(st_row + len(df.index) - 1,
                  st_col + len(df.columns) - 1) # No -1 for the index
         ).Value = df.to_records(index=False)

cmax = len(df.columns)
rmax = len(df.index) + 2

ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax)).Font.Name = "Calibri"
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax)).Font.Size = 11
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax)).Font.Color = 0

ws.Range(ws.Cells(1,1), ws.Cells(1,cmax)).Font.FontStyle = "Bold"

ws.Range("E:E").TextToColumns(Destination=ws.Range("E1"), DataType=c.xlDelimited,TextQualifier=c.xlDoubleQuote, ConsecutiveDelimiter=False, Tab=True, Semicolon=False, Comma=False, Space=False, Other=False,FieldInfo=[1,2],TrailingMinusNumbers=True)
ws.Range("J:J").TextToColumns(Destination=ws.Range("J1"), DataType=c.xlDelimited,TextQualifier=c.xlDoubleQuote, ConsecutiveDelimiter=False, Tab=True, Semicolon=False, Comma=False, Space=False, Other=False,FieldInfo=[1,2],TrailingMinusNumbers=True)
ws.Range("C:C").TextToColumns(Destination=ws.Range("C1"), DataType=c.xlDelimited,TextQualifier=c.xlDoubleQuote, ConsecutiveDelimiter=False, Tab=True, Semicolon=False, Comma=False, Space=False, Other=False,FieldInfo=[1,2],TrailingMinusNumbers=True)

ws.Range("A:N").EntireColumn.AutoFit()


# rs.Close()
wb.SaveAs(resultExcel, c.xlExcel12)
# wb.SaveAs(resultExcel)
cn.close()
wb.Close()
xl.DisplayAlerts = True
xl.Quit()

# ============= CREATE CSV
cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
q = "SELECT tchar FROM stg.Footlocker_Principal_Inventory order by NUM ASC, typ ASC"
curs = cn.cursor()
curs.execute(q)
with open(resultCsv, 'w') as csvfile:
    # writer = csv.writer(csvfile, quoting=csv.QUOTE_NONE)
    # writer.writerow([x[0] for x in curs.description])  # column headers
    for r in curs:
        csvfile.write(str(r).replace('(\'', '').replace('\',)','').replace('(\"', '').replace('\",)',''))
        # print(str(r).replace('(\'', '').replace('\',)',''))
        csvfile.write('\n')
cn.close()
# ===============CREATE sls file
# 2022-03-31 --> Exclude file .sls
# shutil.copyfile(resultCsv, resultSls)

# Output to ssis
# 2022-03-31 --> Exclude file .sls
# print(resultExcel + ";" + resultCsv + ";" + resultSls + ";" + dt_file)
print(resultExcel + ";" + resultCsv + ";" + dt_file)

