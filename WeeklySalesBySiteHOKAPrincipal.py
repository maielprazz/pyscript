import pandas as pd
import pypyodbc as pdb
import win32com.client as mycom
from win32com.client import constants as c
import os
import sys
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

#=============== CREATE EXCEL REPORT ==================
# VARIABLES
workdir = "F:\\03_REPORT\\04_WEEKLY\\Weekly Sales Hoka Principal"
dt_aod = datetime.strptime(str(sys.argv[1]), '%Y%m%d')
dt_st = dt_aod - timedelta(days = dt_aod.weekday())
dt_st_ly = dt_st - relativedelta(years = 1)
dt_end = dt_aod + timedelta(days = 6 - dt_aod.weekday())
dt_end_ly = dt_end - relativedelta(years = 1)
if datetime.strftime(dt_st, '%b') == datetime.strftime(dt_end, '%b') and datetime.strftime(dt_st, '%Y') == datetime.strftime(dt_end, '%Y'):
  dt_file = datetime.strftime(dt_st, '%d') + ' - ' + datetime.strftime(dt_end, '%d') + ' ' + datetime.strftime(dt_end, '%b %Y')
  dt_file_ly = datetime.strftime(dt_st_ly, '%d') + ' - ' + datetime.strftime(dt_end_ly, '%d') + ' ' + datetime.strftime(dt_end_ly, '%b %Y')
elif datetime.strftime(dt_st, '%b') != datetime.strftime(dt_end, '%b') and datetime.strftime(dt_st, '%Y') == datetime.strftime(dt_end, '%Y'):
  dt_file = datetime.strftime(dt_st, '%d %b') + ' - ' + datetime.strftime(dt_end, '%d %b') + ' ' + datetime.strftime(dt_end, '%Y')
  dt_file_ly = datetime.strftime(dt_st_ly, '%d %b') + ' - ' + datetime.strftime(dt_end_ly, '%d %b') + ' ' + datetime.strftime(dt_end_ly, '%Y')
else :
  dt_file = datetime.strftime(dt_st, '%d %b %Y') + ' - ' + datetime.strftime(dt_end, '%d %b %Y')
  dt_file_ly = datetime.strftime(dt_st_ly, '%d %b %Y') + ' - ' + datetime.strftime(dt_end_ly, '%d %b %Y')


xl = mycom.Dispatch('Excel.Application')
# xl = mycom.gencache.EnsureDispatch('Excel.Application')

resultPath = os.path.join(workdir,'WeeklyHOKAReport_' + dt_file.replace(' ', '') +'.xlsx')
# print(resultPath)
xl.Visible = False

wb = xl.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
ws.Name = 'WeeklyHOKA'
xl.ActiveWindow.DisplayGridlines = False
xl.DisplayAlerts = False

cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
q = "EXEC [SP_WEEKLY_HOKA_PRINCIPAL] '0', '" + str(sys.argv[1]) + "'"
df = pd.read_sql(q, cn)

# Paste header
st_row = 1
st_col = 1
for col in df.columns:
  ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
  st_col = st_col + 1

# Paste recordset
st_row = 3
st_col = 1
ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
         ws.Cells(st_row + len(df.index) - 1,
                  st_col + len(df.columns) - 1) # No -1 for the index
         ).Value = df.to_records(index=False)
ws.Range("A1").EntireColumn.Delete()
           

cmax = len(df.columns)
rmax = len(df.index) + 2
# print("rmax", rmax)
# print("cmax", cmax)
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax)).Font.Name = "Calibri"
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax)).Font.Size = 11
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax)).Font.Color = 0

ws.Range(ws.Cells(1,1), ws.Cells(2,cmax)).Font.FontStyle = "Bold"
ws.Range(ws.Cells(1,1), ws.Cells(2,cmax)).HorizontalAlignment = c.xlCenter
ws.Range(ws.Cells(1,1), ws.Cells(2,cmax)).VerticalAlignment = c.xlCenter    

ws.Range("A1:A2").MergeCells = True
ws.Range("B1:B2").MergeCells = True
ws.Range("C1:E1").MergeCells = True
ws.Range("F1:H1").MergeCells = True

ws.Range("A1").Value = "Store Code"      
ws.Range("B1").Value = "Store Name"      
ws.Range("C1").Value = dt_file      
ws.Range("F1").Value = dt_file_ly
   
ws.Range("C2").Value = "Sales Hoka"      
ws.Range("D2").Value = "Traffic"      
ws.Range("E2").Value = "Transaction HOKA Only"
ws.Range("F2").Value = "Sales Hoka"      
ws.Range("G2").Value = "Traffic"      
ws.Range("H2").Value = "Transaction HOKA Only"      

ws.Range("A2").EntireColumn.ColumnWidth = 10
ws.Range("B2").EntireColumn.ColumnWidth = 38
ws.Range("C2").EntireColumn.ColumnWidth = 12
ws.Range("D2").EntireColumn.ColumnWidth = 12
ws.Range("E2").EntireColumn.ColumnWidth = 20
ws.Range("F2").EntireColumn.ColumnWidth = 12
ws.Range("G2").EntireColumn.ColumnWidth = 12
ws.Range("H2").EntireColumn.ColumnWidth = 20

ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.LineStyle = c.xlContinuous
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Weight = c.xlMedium
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Color = 0

ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.LineStyle = c.xlContinuous
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Weight = c.xlMedium
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Color = 0

ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.LineStyle = c.xlContinuous
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Weight = c.xlMedium
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Color = 0

ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.LineStyle = c.xlContinuous
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Weight = c.xlMedium
ws.Range(ws.Cells(1,1), ws.Cells(rmax,cmax-1)).Borders.Color = 0

ws.Range(ws.Cells(3,3), ws.Cells(rmax,cmax-1)).NumberFormatLocal = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"

ws.Range("A1").EntireRow.Insert(c.xlShiftDown, c.xlAbove)
ws.Range("A1").EntireColumn.Insert(c.xlShiftToLeft, c.xlLeft)

# rs.Close()
wb.SaveAs(resultPath)
cn.close()
wb.Close()
xl.DisplayAlerts = True
xl.Quit()

# Output to ssis
print(resultPath + ";" + dt_file)

