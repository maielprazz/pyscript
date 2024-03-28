# to run
# py WeeklySalesFL.py 20220530
from calendar import month
import pandas as pd
# import mysql.connector
import pyodbc
import win32com.client as mycom
from win32com.client import constants as c
import os
import sys
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta

# RGB Function
def rgbtohex(rgb):
  bgr = (rgb[2], rgb[1], rgb[0])
  strValue = '%02x%02x%02x' % bgr
  iValue = int(strValue,16)
  return iValue

#=============== CREATE EXCEL REPORT ==================
# VARIABLES
workdir = "D:\\ISMAIL_ERABW\\reports"
dt_file = "March 2024 - Wellings Daily Sales & Target - DDMMYYYY"

# xl = mycom.Dispatch('Excel.Application')
xl = mycom.gencache.EnsureDispatch('Excel.Application')

resultPath = os.path.join(workdir,'Wellings_Daily_Sales_Target_' + dt_file.replace(' ', '') +'.xlsx')
# print(resultPath)
xl.Visible = False

##########################============ Daily Sales Store REPORT
wb = xl.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
cn = pyodbc.connect("DRIVER={MySQL ODBC 8.3 Unicode Driver}; SERVER=localhost;DATABASE=db_ip; UID=ismail; PASSWORD=Wellings@24;")
#  =============================== SALES TOTAL
q = "SELECT DISTINCT STORE_CODE, STORE_NAME FROM DIM_STORE WHERE STORE_CODE IS NOT NULL order by 1 desc;"
liststore = pd.read_sql(q, cn)
liststore = liststore.reset_index()  # make sure indexes pair with number of rows

for index, row in liststore.iterrows():
    ws = wb.Worksheets.Add()
    ws.Name = row['STORE_CODE']
    
    xl.ActiveWindow.DisplayGridlines = False
    xl.ActiveWindow.Zoom = 70
    xl.DisplayAlerts = False
    
    q = "CALL SP_RPT_DailySalesMTD(1, null, '" + row['STORE_CODE'] + "')"
    dfw = pd.read_sql(q, cn)
    dfw = dfw.fillna('')
    
    ws.Range(ws.Cells(1, 1),ws.Cells(1, 1)).Value = row['STORE_NAME']
    ws.Range(ws.Cells(1, 1),ws.Cells(1, 1)).Font.Size = 16
    ws.Range(ws.Cells(1, 1),ws.Cells(1, 1)).Font.FontStyle = "Bold"
    
    cmax = len(dfw.columns)
    rmax = len(dfw.index) + 4

    # Paste header
    st_row = 4
    st_col = 2
    for col in dfw.columns:
      ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
      st_col = st_col + 1

    ws.Range("B4").Value = "DATE"  
    ws.Range("C4").Value = "Sales"
    ws.Range("D4").Value = "Sales Cumm."
    ws.Range("E4").Value = "Target"
    ws.Range("F4").Value = "Target Cumm."
    ws.Range("G4").Value = "% Ach."
    ws.Range("H4").Value = "% Ach.\n Cumm"
    ws.Range("I4").Value = "Trx"
    ws.Range("J4").Value = "Last \n Month \n Trx"
    ws.Range("K4").Value = "%vs LM"
    ws.Range("L4").Value = "ABV"
    ws.Range("M4").Value = "Last Month \n ABV"
    ws.Range("N4").Value = "%vs LM"
    ws.Range("C3").Value = "TOTAL DAILY SALES"
    ws.Range("C3:N3").MergeCells = True
    ws.Range("B3:B4").MergeCells = True

    # Paste recordset
    st_row = 6
    st_col = 2
    ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
             ws.Cells(st_row + len(dfw.index) - 1,
                      st_col + len(dfw.columns) - 1) # No -1 for the index
             ).Value = dfw.to_records(index=False)
    
    for i in range(6,rmax+2,1): 
      # if row['STORE_CODE'] == 'BT5':
      #   print(i, rmax)
      ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Formula = "=DATE(LEFT(B" + str(i) + ",4),MID(B" + str(i) + ",5,2),RIGHT(B" + str(i) + ",2))"

    ws.Range(ws.Cells(6,1), ws.Cells(rmax+1,1)).Copy()
    ws.Range(ws.Cells(6,1), ws.Cells(6,1)).PasteSpecial(Paste=c.xlPasteValues) 
    ws.Range(ws.Cells(6,1), ws.Cells(rmax+1,1)).Cut(ws.Range(ws.Cells(6,2), ws.Cells(6,2)))

    ws.Range("C:N").NumberFormat = "#,#0"
    ws.Range("G:H").NumberFormat = "0.0%"
    ws.Range("K:K").NumberFormat = "0.0%"
    ws.Range("N:N").NumberFormat = "0.0%"
    
    # Title Font
    ws.Range(ws.Cells(1,2), ws.Cells(1,2)).Font.FontStyle = "Bold"

    # Header Format
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Font.Name = "Calibri"
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Font.FontStyle = "Bold"
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Font.Size = 11
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Font.Color = rgbtohex((255,255,255))
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Interior.Color = rgbtohex((0,20,60))
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).HorizontalAlignment = c.xlCenter
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).VerticalAlignment = c.xlCenter
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Borders.Color = 0
    ws.Range("C4:D4").Interior.Color = rgbtohex((0,112,192))
    
    # body border
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders.Weight = c.xlHairline
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders.Color = 0

    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(4,cmax+1)).Borders.Color = 0

    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlInsideVertical).Color = 0

    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeTop).Color = 0
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeBottom).Color = 0
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeLeft).Color = 0
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
    ws.Range(ws.Cells(3,2), ws.Cells(rmax+2,cmax+1)).Borders(c.xlEdgeRight).Color = 0

    #grand total
    ws.Range(ws.Cells(rmax+2,2), ws.Cells(rmax+2,cmax+1)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(rmax+2,2), ws.Cells(rmax+2,cmax+1)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(rmax+2,2), ws.Cells(rmax+2,cmax+1)).Borders.Color = 0
    ws.Range(ws.Cells(rmax+2,2), ws.Cells(rmax+2,cmax+1)).Font.FontStyle = "Bold"

    # row color
    ws.Range(ws.Cells(5,2), ws.Cells(5,cmax+1)).Interior.Color = rgbtohex((192,0,0))

    clor = [7,8,14,15,21,22,28,29,35,36]
    clorset = set(clor)
    for ln in range(1, rmax+2,1):
      if ln in clorset: 
        ws.Range(ws.Cells(ln,2), ws.Cells(ln,cmax+1)).Interior.Color = rgbtohex((253,233,217))

    ws.Range(ws.Cells(rmax+2,2), ws.Cells(rmax+2,cmax+1)).Interior.Color = rgbtohex((196,189,151))

    # width and height
    ws.Columns("A:A").ColumnWidth = 0.44
    ws.Columns("B:B").ColumnWidth = 9.67
    ws.Columns("C:E").ColumnWidth = 11.44
    ws.Columns("F:F").ColumnWidth = 12.56
    ws.Columns("G:H").ColumnWidth = 6.67
    ws.Columns("I:J").ColumnWidth = 8.78
    ws.Columns("K:N").ColumnWidth = 11.44
    ws.Rows(4).RowHeight = 43.8

    # # Formula
    # grand total
    ws.Range(ws.Cells(rmax+2,3), ws.Cells(rmax+2,3)).Formula = "=SUM(C6:C" + str(rmax+1) + ")"
    ws.Range(ws.Cells(rmax+2,4), ws.Cells(rmax+2,4)).Formula = "=+D" + str(rmax+1) + ""
    ws.Range(ws.Cells(rmax+2,5), ws.Cells(rmax+2,5)).Formula = "=SUM(E6:E" + str(rmax+1) + ")"
    ws.Range(ws.Cells(rmax+2,6), ws.Cells(rmax+2,6)).Formula = "=+F" + str(rmax+1) + ""
    ws.Range(ws.Cells(rmax+2,7), ws.Cells(rmax+2,7)).Formula = "=IFERROR(C" + str(rmax+2) + "/E"+ str(rmax+2) + ",0)" 
    ws.Range(ws.Cells(rmax+2,8), ws.Cells(rmax+2,8)).Formula = "=IFERROR(D" + str(rmax+2) + "/F"+ str(rmax+2) + ",0)" 
    ws.Range(ws.Cells(rmax+2,9), ws.Cells(rmax+2,9)).Formula = "=SUM(I6:I" + str(rmax+1) + ")"
    ws.Range(ws.Cells(rmax+2,10), ws.Cells(rmax+2,10)).Formula = "=SUM(J6:J" + str(rmax+1) + ")"
    ws.Range(ws.Cells(rmax+2,11), ws.Cells(rmax+2,11)).Formula = "=IFERROR(I" + str(rmax+2) + "/J"+ str(rmax+2) + "-1,0)"
    ws.Range(ws.Cells(rmax+2,12), ws.Cells(rmax+2,12)).Formula = "=IFERROR(C" + str(rmax+2) + "/I"+ str(rmax+2) + ",0)"  
    ws.Range(ws.Cells(rmax+2,13), ws.Cells(rmax+2,13)).Formula = "=IFERROR(D" + str(rmax+2) + "/J"+ str(rmax+2) + ",0)"
    ws.Range(ws.Cells(rmax+2,14), ws.Cells(rmax+2,14)).Formula = "=IFERROR(L" + str(rmax+2) + "/M"+ str(rmax+2) + "-1,0)"

    # inside table
    if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
      rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
    else:
      rmaxdata = rmax + 1

    ws.Range("D6").Formula = "=C6"
    ws.Range("F6").Formula = "=E6"
    for li in range(7,rmax+2,1):
       ws.Range(ws.Cells(li,4), ws.Cells(li,4)).Formula = "=D" + str(li-1) + "+C" + str(li) +""
       ws.Range(ws.Cells(li,6), ws.Cells(li,6)).Formula = "=F" + str(li-1) + "+E" + str(li) +"" 
    
    for il in range(6,rmaxdata+1,1):   
      ws.Range(ws.Cells(il,7), ws.Cells(il,7)).Formula = "=IFERROR(C" + str(il) + "/E"+ str(il) + ",0)" 
      ws.Range(ws.Cells(il,8), ws.Cells(il,8)).Formula = "=IFERROR(D" + str(il) + "/F"+ str(il) + ",0)" 
      ws.Range(ws.Cells(il,11), ws.Cells(il,11)).Formula = "=IFERROR(I" + str(il) + "/J"+ str(il) + "-1,0)"
      ws.Range(ws.Cells(il,12), ws.Cells(il,12)).Formula = "=IFERROR(C" + str(il) + "/I"+ str(il) + "-1,0)"
      ws.Range(ws.Cells(il,14), ws.Cells(il,14)).Formula = "=IFERROR(L" + str(il) + "/M"+ str(il) + "-1,0)"

# #  =============== HALODOC  =======================================================================

#     q = "CALL SP_RPT_DailySalesMTD(2, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 4
#     st_col = 16
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1

#     ws.Range("P4").Value = "DATE"  
#     ws.Range("Q4").Value = "Sales"
#     ws.Range("R4").Value = "Sales Cumm."
#     ws.Range("S4").Value = "Target"
#     ws.Range("T4").Value = "Target Cumm."
#     ws.Range("U4").Value = "% Ach."
#     ws.Range("V4").Value = "% Ach.\n Cumm"
#     ws.Range("Q3").Value = "HALODOC"
#     ws.Range("Q3:V3").MergeCells = True
#     ws.Range("P3:P4").MergeCells = True

#     # Paste recordset
#     st_row = 6
#     st_col = 16
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(6,rmax+2,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(P" + str(i) + ",4),MID(P" + str(i) + ",5,2),RIGHT(P" + str(i) + ",2))"

#     ws.Range(ws.Cells(6,st_col-1), ws.Cells(rmax+1,st_col-1)).Copy()
#     ws.Range(ws.Cells(6,st_col-1), ws.Cells(6,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(6,st_col-1), ws.Cells(rmax+1,st_col-1)).Cut(ws.Range(ws.Cells(6,st_col), ws.Cells(6,st_col)))

#     ws.Range("P" + str(st_row) + ":V" + str(rmax+2) + "").NumberFormat = "#,#0"
#     ws.Range("U" + str(st_row) + ":V" + str(rmax+2) + "").NumberFormat = "0.0%"
#     ws.Range("P" + str(st_row) + ":P" + str(rmax+1) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range("Q4:R4").Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(5,st_col), ws.Cells(5,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(ln,st_col), ws.Cells(ln,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # width and height
#     ws.Columns("P:P").ColumnWidth = 9.67
#     ws.Columns("R:R").ColumnWidth = 8.78
#     ws.Columns("S:S").ColumnWidth = 9.33
#     ws.Columns("T:T").ColumnWidth = 11.44
#     ws.Columns("U:U").ColumnWidth = 5.56
#     ws.Columns("V:V").ColumnWidth = 11.44

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(rmax+2,st_col+1), ws.Cells(rmax+2,st_col+1)).Formula = "=SUM(Q6:Q" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+2), ws.Cells(rmax+2,st_col+2)).Formula = "=+R" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+3), ws.Cells(rmax+2,st_col+3)).Formula = "=SUM(S6:S" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+4), ws.Cells(rmax+2,st_col+4)).Formula = "=+T" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+5), ws.Cells(rmax+2,st_col+5)).Formula ="=IFERROR(Q" + str(rmax+2) + "/S"+ str(rmax+2) + ",0)" 
#     ws.Range(ws.Cells(rmax+2,st_col+6), ws.Cells(rmax+2,st_col+6)).Formula ="=IFERROR(R" + str(rmax+2) + "/T"+ str(rmax+2) + ",0)" 
    
#     # inside table
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1
#     ws.Range("R6").Formula = "=Q6"
#     ws.Range("T6").Formula = "=S6"
    
#     for li in range(6,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(Q" + str(li) + "/S"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(R" + str(li) + "/T"+ str(li) + ",0)" 

#     for li in range(7,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=R" + str(li-1) + "+Q" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=T" + str(li-1) + "+S" + str(li) +""    

# #  =============== CEK KES  =======================================================================

#     q = "CALL SP_RPT_DailySalesMTD(3, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 4
#     st_col = 24
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
    
#     st_col = 24
#     # if row['STORE_CODE'] == 'BT5':
#     #   print(st_row, st_col)
    
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Sales"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Sales Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "CEKKES"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 6
#     st_col = 24
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,rmax+2,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(X" + str(i) + ",4),MID(X" + str(i) + ",5,2),RIGHT(X" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("Y" + str(st_row) + ":AD" + str(rmax+2) + "").NumberFormat = "#,#0"
#     ws.Range("AC" + str(st_row) + ":AD" + str(rmax+2) + "").NumberFormat = "0.0%"
#     ws.Range("X" + str(st_row) + ":X" + str(rmax+1) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(ln,st_col), ws.Cells(ln,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # width and height
#     # ws.Union(Columns(st_col)).ColumnWidth = 9.67
#     ws.Columns(st_col).ColumnWidth = 10
#     ws.Columns(st_col+1).ColumnWidth = 12
#     ws.Columns(st_col+2).ColumnWidth = 12
#     ws.Columns(st_col+3).ColumnWidth = 12
#     ws.Columns(st_col+4).ColumnWidth = 13
#     ws.Columns(st_col+5).ColumnWidth = 7
#     ws.Columns(st_col+6).ColumnWidth = 7

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(rmax+2,st_col+1), ws.Cells(rmax+2,st_col+1)).Formula = "=SUM(Y6:Y" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+2), ws.Cells(rmax+2,st_col+2)).Formula = "=+Z" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+3), ws.Cells(rmax+2,st_col+3)).Formula = "=SUM(AA6:AA" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+4), ws.Cells(rmax+2,st_col+4)).Formula = "=+AB" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+5), ws.Cells(rmax+2,st_col+5)).Formula ="=IFERROR(Y" + str(rmax+2) + "/AA"+ str(rmax+2) + ",0)" 
#     ws.Range(ws.Cells(rmax+2,st_col+6), ws.Cells(rmax+2,st_col+6)).Formula ="=IFERROR(Z" + str(rmax+2) + "/AB"+ str(rmax+2) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=Y6"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=AA6"
    
#     for li in range(st_row,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(Y" + str(li) + "/AA"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(Z" + str(li) + "/AB"+ str(li) + ",0)" 

#     for li in range(st_row+1,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=Z" + str(li-1) + "+Y" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=AB" + str(li-1) + "+AA" + str(li) +""    

# #  =============== PWP  =======================================================================

#     q = "CALL SP_RPT_DailySalesMTD(4, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 4
#     st_col = 32
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 32
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Sales"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Sales Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "PWP"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 6
#     st_col = 32
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,rmax+2,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(AF" + str(i) + ",4),MID(AF" + str(i) + ",5,2),RIGHT(AF" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("AG" + str(st_row) + ":AL" + str(rmax+2) + "").NumberFormat = "#,#0"
#     ws.Range("AK" + str(st_row) + ":AL" + str(rmax+2) + "").NumberFormat = "0.0%"
#     ws.Range("AF" + str(st_row) + ":AF" + str(rmax+1) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(ln,st_col), ws.Cells(ln,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # width and height
#     # ws.Union(Columns(st_col)).ColumnWidth = 9.67
#     ws.Columns(st_col).ColumnWidth = 10
#     ws.Columns(st_col+1).ColumnWidth = 12
#     ws.Columns(st_col+2).ColumnWidth = 12
#     ws.Columns(st_col+3).ColumnWidth = 12
#     ws.Columns(st_col+4).ColumnWidth = 13
#     ws.Columns(st_col+5).ColumnWidth = 7
#     ws.Columns(st_col+6).ColumnWidth = 7


#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(rmax+2,st_col+1), ws.Cells(rmax+2,st_col+1)).Formula = "=SUM(AG6:AG" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+2), ws.Cells(rmax+2,st_col+2)).Formula = "=+AH" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+3), ws.Cells(rmax+2,st_col+3)).Formula = "=SUM(AI6:AI" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+4), ws.Cells(rmax+2,st_col+4)).Formula = "=+AJ" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+5), ws.Cells(rmax+2,st_col+5)).Formula ="=IFERROR(AG" + str(rmax+2) + "/AI"+ str(rmax+2) + ",0)" 
#     ws.Range(ws.Cells(rmax+2,st_col+6), ws.Cells(rmax+2,st_col+6)).Formula ="=IFERROR(AH" + str(rmax+2) + "/AJ"+ str(rmax+2) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=AG6"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=AI6"
    
#     for li in range(st_row,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(AG" + str(li) + "/AI"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(AH" + str(li) + "/AJ"+ str(li) + ",0)" 

#     for li in range(st_row+1,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=AH" + str(li-1) + "+AG" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=AJ" + str(li-1) + "+AI" + str(li) +""    

# #  =============== THematic =======================================================================

#     q = "CALL SP_RPT_DailySalesMTD(5, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 4
#     st_col = 40
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 40
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Sales"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Sales Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "THEMATIC"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 6
#     st_col = 40
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,rmax+2,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(AN" + str(i) + ",4),MID(AN" + str(i) + ",5,2),RIGHT(AN" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("AO" + str(st_row) + ":AT" + str(rmax+2) + "").NumberFormat = "#,#0"
#     ws.Range("AS" + str(st_row) + ":AT" + str(rmax+2) + "").NumberFormat = "0.0%"
#     ws.Range("AN" + str(st_row) + ":AN" + str(rmax+1) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(ln,st_col), ws.Cells(ln,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # width and height
#     # ws.Union(Columns(st_col)).ColumnWidth = 9.67
#     ws.Columns(st_col).ColumnWidth = 10
#     ws.Columns(st_col+1).ColumnWidth = 12
#     ws.Columns(st_col+2).ColumnWidth = 12
#     ws.Columns(st_col+3).ColumnWidth = 12
#     ws.Columns(st_col+4).ColumnWidth = 13
#     ws.Columns(st_col+5).ColumnWidth = 7
#     ws.Columns(st_col+6).ColumnWidth = 7


#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(rmax+2,st_col+1), ws.Cells(rmax+2,st_col+1)).Formula = "=SUM(AO6:AO" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+2), ws.Cells(rmax+2,st_col+2)).Formula = "=+AP" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+3), ws.Cells(rmax+2,st_col+3)).Formula = "=SUM(AQ6:AQ" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+4), ws.Cells(rmax+2,st_col+4)).Formula = "=+AR" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+5), ws.Cells(rmax+2,st_col+5)).Formula ="=IFERROR(AO" + str(rmax+2) + "/AQ"+ str(rmax+2) + ",0)" 
#     ws.Range(ws.Cells(rmax+2,st_col+6), ws.Cells(rmax+2,st_col+6)).Formula ="=IFERROR(AP" + str(rmax+2) + "/AR"+ str(rmax+2) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=AO6"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=AQ6"
    
#     for li in range(st_row,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(AO" + str(li) + "/AQ"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(AP" + str(li) + "/AR"+ str(li) + ",0)" 

#     for li in range(st_row+1,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=AP" + str(li-1) + "+AO" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=AR" + str(li-1) + "+AQ" + str(li) +""    

#     #  =============== Best Choice =======================================================================

#     q = "CALL SP_RPT_DailySalesMTD(6, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 4
#     st_col = 48
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 48
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Sales"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Sales Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "BEST CHOICE"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 6
#     st_col = 48
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,rmax+2,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(AV" + str(i) + ",4),MID(AV" + str(i) + ",5,2),RIGHT(AV" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("AW" + str(st_row) + ":BB" + str(rmax+2) + "").NumberFormat = "#,#0"
#     ws.Range("BA" + str(st_row) + ":BB" + str(rmax+2) + "").NumberFormat = "0.0%"
#     ws.Range("AV" + str(st_row) + ":AV" + str(rmax+1) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(ln,st_col), ws.Cells(ln,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # width and height
#     # ws.Union(Columns(st_col)).ColumnWidth = 9.67
#     ws.Columns(st_col).ColumnWidth = 10
#     ws.Columns(st_col+1).ColumnWidth = 12
#     ws.Columns(st_col+2).ColumnWidth = 12
#     ws.Columns(st_col+3).ColumnWidth = 12
#     ws.Columns(st_col+4).ColumnWidth = 13
#     ws.Columns(st_col+5).ColumnWidth = 7
#     ws.Columns(st_col+6).ColumnWidth = 7


#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(rmax+2,st_col+1), ws.Cells(rmax+2,st_col+1)).Formula = "=SUM(AW6:AW" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+2), ws.Cells(rmax+2,st_col+2)).Formula = "=+AX" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+3), ws.Cells(rmax+2,st_col+3)).Formula = "=SUM(AY6:AY" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+4), ws.Cells(rmax+2,st_col+4)).Formula = "=+AZ" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+5), ws.Cells(rmax+2,st_col+5)).Formula ="=IFERROR(AW" + str(rmax+2) + "/AY"+ str(rmax+2) + ",0)" 
#     ws.Range(ws.Cells(rmax+2,st_col+6), ws.Cells(rmax+2,st_col+6)).Formula ="=IFERROR(AX" + str(rmax+2) + "/AZ"+ str(rmax+2) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=AW6"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=AY6"
    
#     for li in range(st_row,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(AW" + str(li) + "/AY"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(AX" + str(li) + "/AZ"+ str(li) + ",0)" 

#     for li in range(st_row+1,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=AX" + str(li-1) + "+AW" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=AZ" + str(li-1) + "+AY" + str(li) +""    

# #  =============== Pharma =======================================================================

#     q = "CALL SP_RPT_DailySalesMTD(7, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 4
#     st_col = 56
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 56
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Sales"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Sales Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "PHARMA"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 6
#     st_col = 56
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,rmax+2,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(BD" + str(i) + ",4),MID(BD" + str(i) + ",5,2),RIGHT(BD" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(rmax+1,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("BE" + str(st_row) + ":BJ" + str(rmax+2) + "").NumberFormat = "#,#0"
#     ws.Range("BI" + str(st_row) + ":BJ" + str(rmax+2) + "").NumberFormat = "0.0%"
#     ws.Range("BD" + str(st_row) + ":BD" + str(rmax+1) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(ln,st_col), ws.Cells(ln,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(rmax+2,st_col), ws.Cells(rmax+2,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # width and height
#     # ws.Union(Columns(st_col)).ColumnWidth = 9.67
#     ws.Columns(st_col).ColumnWidth = 10
#     ws.Columns(st_col+1).ColumnWidth = 12
#     ws.Columns(st_col+2).ColumnWidth = 12
#     ws.Columns(st_col+3).ColumnWidth = 12
#     ws.Columns(st_col+4).ColumnWidth = 13
#     ws.Columns(st_col+5).ColumnWidth = 7
#     ws.Columns(st_col+6).ColumnWidth = 7

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(rmax+2,st_col+1), ws.Cells(rmax+2,st_col+1)).Formula = "=SUM(BE6:BE" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+2), ws.Cells(rmax+2,st_col+2)).Formula = "=+BF" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+3), ws.Cells(rmax+2,st_col+3)).Formula = "=SUM(BG6:BG" + str(rmax+1) + ")"
#     ws.Range(ws.Cells(rmax+2,st_col+4), ws.Cells(rmax+2,st_col+4)).Formula = "=+BH" + str(rmax+1) + ""
#     ws.Range(ws.Cells(rmax+2,st_col+5), ws.Cells(rmax+2,st_col+5)).Formula ="=IFERROR(BE" + str(rmax+2) + "/BG"+ str(rmax+2) + ",0)" 
#     ws.Range(ws.Cells(rmax+2,st_col+6), ws.Cells(rmax+2,st_col+6)).Formula ="=IFERROR(BF" + str(rmax+2) + "/BH"+ str(rmax+2) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=BE6"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=BG6"
    
#     for li in range(st_row,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(AW" + str(li) + "/AY"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(AX" + str(li) + "/AZ"+ str(li) + ",0)" 

#     for li in range(st_row+1,rmax+2,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=BF" + str(li-1) + "+BE" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=BH" + str(li-1) + "+BG" + str(li) +""    

# #  =============== No Of Existing Member =======================================================================
#     q = "CALL SP_RPT_DailySalesMTD(8, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 41
#     st_col = 2
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 2
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Actual"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Actual Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "NO OF EXISTING MEMBER"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 43
#     st_col = 2
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,st_row+rmax-4,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(B" + str(i) + ",4),MID(B" + str(i) + ",5,2),RIGHT(B" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("C" + str(st_row) + ":H" + str(st_row+rmax-4) + "").NumberFormat = "#,#0"
#     ws.Range("G" + str(st_row) + ":H" + str(st_row+rmax-4) + "").NumberFormat = "0.0%"
#     ws.Range("B" + str(st_row) + ":B" + str(st_row+rmax-5) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(st_row+ln-6,st_col), ws.Cells(st_row+ln-6,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+1), ws.Cells(st_row+rmax-4,st_col+1)).Formula = "=SUM(C43:C" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+2), ws.Cells(st_row+rmax-4,st_col+2)).Formula = "=+D" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+3), ws.Cells(st_row+rmax-4,st_col+3)).Formula = "=SUM(E43:E" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+4), ws.Cells(st_row+rmax-4,st_col+4)).Formula = "=+F" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+5), ws.Cells(st_row+rmax-4,st_col+5)).Formula ="=IFERROR(C" + str(st_row+rmax-4) + "/E"+ str(st_row+rmax-4) + ",0)" 
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+6), ws.Cells(st_row+rmax-4,st_col+6)).Formula ="=IFERROR(D" + str(st_row+rmax-4) + "/F"+ str(st_row+rmax-4) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=C43"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=E43"
    
#     for li in range(st_row,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(C" + str(li) + "/E"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(D" + str(li) + "/F"+ str(li) + ",0)" 

#     for li in range(st_row+1,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=D" + str(li-1) + "+C" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=F" + str(li-1) + "+E" + str(li) +""    

# #  =============== TRX Existing Member =======================================================================
#     q = "CALL SP_RPT_DailySalesMTD(9, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 41
#     st_col = 10
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 10
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Actual"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Actual Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "TRANSACTION EXISTING MEMBER"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 43
#     st_col = 10
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,st_row+rmax-4,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(J" + str(i) + ",4),MID(J" + str(i) + ",5,2),RIGHT(J" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("K" + str(st_row) + ":P" + str(st_row+rmax-4) + "").NumberFormat = "#,#0"
#     ws.Range("O" + str(st_row) + ":P" + str(st_row+rmax-4) + "").NumberFormat = "0.0%"
#     ws.Range("J" + str(st_row) + ":J" + str(st_row+rmax-5) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(st_row+ln-6,st_col), ws.Cells(st_row+ln-6,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+1), ws.Cells(st_row+rmax-4,st_col+1)).Formula = "=SUM(K43:K" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+2), ws.Cells(st_row+rmax-4,st_col+2)).Formula = "=+L" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+3), ws.Cells(st_row+rmax-4,st_col+3)).Formula = "=SUM(M43:M" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+4), ws.Cells(st_row+rmax-4,st_col+4)).Formula = "=+N" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+5), ws.Cells(st_row+rmax-4,st_col+5)).Formula ="=IFERROR(K" + str(st_row+rmax-4) + "/M"+ str(st_row+rmax-4) + ",0)" 
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+6), ws.Cells(st_row+rmax-4,st_col+6)).Formula ="=IFERROR(L" + str(st_row+rmax-4) + "/N"+ str(st_row+rmax-4) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=K43"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=M43"
    
#     for li in range(st_row,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(K" + str(li) + "/M"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(L" + str(li) + "/N"+ str(li) + ",0)" 

#     for li in range(st_row+1,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=L" + str(li-1) + "+K" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=N" + str(li-1) + "+M" + str(li) +""    

# #  =============== No Of New Member =======================================================================
#     q = "CALL SP_RPT_DailySalesMTD(10, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 41
#     st_col = 18
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 18
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Actual"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Actual Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "NO OF NEW MEMBER"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 43
#     st_col = 18
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,st_row+rmax-4,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(R" + str(i) + ",4),MID(R" + str(i) + ",5,2),RIGHT(R" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("S" + str(st_row) + ":X" + str(st_row+rmax-4) + "").NumberFormat = "#,#0"
#     ws.Range("W" + str(st_row) + ":X" + str(st_row+rmax-4) + "").NumberFormat = "0.0%"
#     ws.Range("R" + str(st_row) + ":R" + str(st_row+rmax-5) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(st_row+ln-6,st_col), ws.Cells(st_row+ln-6,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+1), ws.Cells(st_row+rmax-4,st_col+1)).Formula = "=SUM(S43:S" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+2), ws.Cells(st_row+rmax-4,st_col+2)).Formula = "=+T" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+3), ws.Cells(st_row+rmax-4,st_col+3)).Formula = "=SUM(U43:U" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+4), ws.Cells(st_row+rmax-4,st_col+4)).Formula = "=+V" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+5), ws.Cells(st_row+rmax-4,st_col+5)).Formula ="=IFERROR(S" + str(st_row+rmax-4) + "/U"+ str(st_row+rmax-4) + ",0)" 
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+6), ws.Cells(st_row+rmax-4,st_col+6)).Formula ="=IFERROR(T" + str(st_row+rmax-4) + "/V"+ str(st_row+rmax-4) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=S43"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=U43"
    
#     for li in range(st_row,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(S" + str(li) + "/U"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(T" + str(li) + "/V"+ str(li) + ",0)" 

#     for li in range(st_row+1,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=T" + str(li-1) + "+S" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=V" + str(li-1) + "+U" + str(li) +""    

# #  =============== trx Of New+Non Member =======================================================================
#     q = "CALL SP_RPT_DailySalesMTD(11, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 41
#     st_col = 26
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 26
#     ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "DATE"  
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Actual"
#     ws.Range(ws.Cells(st_row, st_col+2),ws.Cells(st_row, st_col+2)).Value = "Actual Cumm."
#     ws.Range(ws.Cells(st_row, st_col+3),ws.Cells(st_row, st_col+3)).Value = "Target"
#     ws.Range(ws.Cells(st_row, st_col+4),ws.Cells(st_row, st_col+4)).Value = "Target Cumm."
#     ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "% Ach."
#     ws.Range(ws.Cells(st_row, st_col+6),ws.Cells(st_row, st_col+6)).Value = "% Ach.\n Cumm"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+1)).Value = "TRANSACTION OF NEW+NON MEMBER"
#     ws.Range(ws.Cells(st_row-1, st_col+1),ws.Cells(st_row-1, st_col+cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row, st_col)).MergeCells = True

#     # Paste recordset
#     st_row = 43
#     st_col = 26
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     for i in range(st_row,st_row+rmax-4,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(Z" + str(i) + ",4),MID(Z" + str(i) + ",5,2),RIGHT(Z" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))

#     ws.Range("AA" + str(st_row) + ":AF" + str(st_row+rmax-4) + "").NumberFormat = "#,#0"
#     ws.Range("AE" + str(st_row) + ":AF" + str(st_row+rmax-4) + "").NumberFormat = "0.0%"
#     ws.Range("Z" + str(st_row) + ":Z" + str(st_row+rmax-5) + "").NumberFormat = "m/d/yyyy"

#     # Header Format
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Font.Color = rgbtohex((255,255,255))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Interior.Color = rgbtohex((0,20,60))
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row-2,st_col+1), ws.Cells(st_row-2,st_col+2)).Interior.Color = rgbtohex((0,112,192))
    
#     # body border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-2,st_col-2+cmax+1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders(c.xlEdgeRight).Color = 0

#     #grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Borders.Color = 0
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Font.FontStyle = "Bold"

#     # row color
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col-2+cmax+1)).Interior.Color = rgbtohex((192,0,0))

#     clor = [7,8,14,15,21,22,28,29,35,36]
#     clorset = set(clor)

#     for ln in range(1, rmax+2,1):
#       if ln in clorset: 
#         ws.Range(ws.Cells(st_row+ln-6,st_col), ws.Cells(st_row+ln-6,st_col-2+cmax+1)).Interior.Color = rgbtohex((253,233,217))

#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col-2+cmax+1)).Interior.Color = rgbtohex((196,189,151))

#     # # Formula
#     # grand total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+1), ws.Cells(st_row+rmax-4,st_col+1)).Formula = "=SUM(AA43:AA" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+2), ws.Cells(st_row+rmax-4,st_col+2)).Formula = "=+AB" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+3), ws.Cells(st_row+rmax-4,st_col+3)).Formula = "=SUM(AC43:AC" + str(st_row+rmax-5) + ")"
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+4), ws.Cells(st_row+rmax-4,st_col+4)).Formula = "=+AD" + str(st_row+rmax-5) + ""
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+5), ws.Cells(st_row+rmax-4,st_col+5)).Formula ="=IFERROR(AA" + str(st_row+rmax-4) + "/AC"+ str(st_row+rmax-4) + ",0)" 
#     ws.Range(ws.Cells(st_row+rmax-4,st_col+6), ws.Cells(st_row+rmax-4,st_col+6)).Formula ="=IFERROR(AB" + str(st_row+rmax-4) + "/AD"+ str(st_row+rmax-4) + ",0)" 
    
#     # inside table cek only C3
#     if ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).Value is None:
#       rmaxdata = ws.Range(ws.Cells(rmax+1,3), ws.Cells(rmax+1,3)).End(c.xlUp).Row
#     else:
#       rmaxdata = rmax + 1

#     ws.Range(ws.Cells(st_row,st_col+2), ws.Cells(st_row, st_col+2)).Formula = "=AA43"
#     ws.Range(ws.Cells(st_row,st_col+4), ws.Cells(st_row, st_col+4)).Formula = "=AC43"
    
#     for li in range(st_row,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+5), ws.Cells(li,st_col+5)).Formula = "=IFERROR(AA" + str(li) + "/AC"+ str(li) + ",0)" 
#        ws.Range(ws.Cells(li,st_col+6), ws.Cells(li,st_col+6)).Formula = "=IFERROR(AB" + str(li) + "/AD"+ str(li) + ",0)" 

#     for li in range(st_row+1,st_row+rmax-4,1):
#        ws.Range(ws.Cells(li,st_col+2), ws.Cells(li,st_col+2)).Formula = "=AB" + str(li-1) + "+AA" + str(li) +""
#        ws.Range(ws.Cells(li,st_col+4), ws.Cells(li,st_col+4)).Formula = "=AD" + str(li-1) + "+AC" + str(li) +""    
    
#  =============== TOP 30 SKU =======================================================================
    q = "CALL SP_RPT_DailySalesMTD(12, null, '" + row['STORE_CODE'] + "')"
    dfw = pd.read_sql(q, cn)
    dfw = dfw.fillna('')
    
    cmax = len(dfw.columns)
    rmax = len(dfw.index) + 4

    # Paste header
    st_row = 41
    st_col = 34
    # for col in dfw.columns:
    #   ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
    #   st_col = st_col + 1
    # st_col = 34
    
    ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = "Rank"  
    ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+1)).Value = "Item Name"
    ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+5)).Value = "Sales"
    ws.Range(ws.Cells(st_row, st_col+7),ws.Cells(st_row, st_col+7)).Value = "Transaction"
    ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row-1, st_col)).Formula = "=\"MTD \" & UPPER(TEXT(B6,\"MMMM\"))"
    ws.Range(ws.Cells(st_row-2, st_col),ws.Cells(st_row-2,st_col)).Value = "TOP 30 SKUs"
    
  # Header Format
    ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row-1, st_col+cmax)).MergeCells = True
    ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row,st_col+cmax)).Font.Name = "Calibri"
    ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row,st_col+cmax)).Font.FontStyle = "Bold"
    ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row,st_col+cmax)).Font.Size = 11
    ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row,st_col+cmax)).Font.Color = rgbtohex((255,255,255)) # putih
    ws.Range(ws.Cells(st_row-2, st_col),ws.Cells(st_row-2,st_col+1)).Interior.Color = rgbtohex((255,255,0)) # kuning
    ws.Range(ws.Cells(st_row-1, st_col),ws.Cells(st_row-1, st_col+cmax)).Interior.Color = rgbtohex((0,20,60)) #dark
    ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col+cmax)).Interior.Color = rgbtohex((0,176,240)) # light blue
    ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row, st_col+4)).MergeCells = True
    ws.Range(ws.Cells(st_row, st_col+5),ws.Cells(st_row, st_col+6)).MergeCells = True

    ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row,st_col+cmax)).HorizontalAlignment = c.xlCenter
    ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row,st_col+cmax)).VerticalAlignment = c.xlCenter
    ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row,st_col+cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row,st_col+cmax)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row,st_col+cmax)).Borders.Color = 0
    
    # Paste recordset
    st_row = 42
    st_col = 35
    ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
             ws.Cells(st_row + len(dfw.index) - 1,
                      st_col + len(dfw.columns) - 1) # No -1 for the index
             ).Value = dfw.to_records(index=False)
    ws.Range("AM" + str(st_row) + ":AO" + str(st_row+rmax-4) + "").NumberFormat = "#,#0"
    st_col = 34
    # numbering
    ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)).Formula = "=1" 
    for li in range(st_row+1,st_row+rmax-4,1):
       ws.Range(ws.Cells(li,st_col), ws.Cells(li,st_col)).Formula = "=AH" + str(li-1) + "+1" 
    
    # Merge cells in body
    for j in range(st_row,st_row+rmax,1):
      ws.Range(ws.Cells(j, st_col+1),ws.Cells(j, st_col+4)).MergeCells = True
      ws.Range(ws.Cells(j, st_col+5),ws.Cells(j, st_col+6)).MergeCells = True

    # # body border
    # if row['STORE_CODE'] == 'VTR' :
    #   print(rmax, cmax) # 34 7
      

    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders.Weight = c.xlHairline
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders.Color = 0

    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row-2,st_col+cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row-2,st_col+cmax)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row-2,st_col+cmax)).Borders.Color = 0

    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlInsideVertical).Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlInsideVertical).Color = 0

    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeTop).Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeTop).Color = 0
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeBottom).Color = 0
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeLeft).Color = 0
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeRight).Weight = c.xlMedium
    ws.Range(ws.Cells(st_row-2,st_col), ws.Cells(st_row+rmax-5,st_col+cmax)).Borders(c.xlEdgeRight).Color = 0

# #  =============== CEK KES by Store by Person =======================================================================
#     q = "CALL SP_RPT_DailySalesMTD(13, null, '" + row['STORE_CODE'] + "')"
#     dfw = pd.read_sql(q, cn)
#     dfw = dfw.fillna('')
    
#     cmax = len(dfw.columns)
#     rmax = len(dfw.index) + 4

#     # Paste header
#     st_row = 82
#     st_col = 2
#     for col in dfw.columns:
#       ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
#       st_col = st_col + 1
#     st_col = 2

#     # Paste recordset
#     st_row = 83
#     st_col = 2
#     ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
#              ws.Cells(st_row + len(dfw.index) - 1,
#                       st_col + len(dfw.columns) - 1) # No -1 for the index
#              ).Value = dfw.to_records(index=False)
    
#     st_col = 2
    
#     for i in range(st_row,st_row+rmax-4,1): 
#       ws.Range(ws.Cells(i,st_col-1), ws.Cells(i,st_col-1)).Formula = "=DATE(LEFT(B" + str(i) + ",4),MID(B" + str(i) + ",5,2),RIGHT(B" + str(i) + ",2))"

#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Copy()
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row,st_col-1)).PasteSpecial(Paste=c.xlPasteValues) 
#     ws.Range(ws.Cells(st_row,st_col-1), ws.Cells(st_row+rmax-5,st_col-1)).Cut(ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row,st_col)))
#     ws.Range("B" + str(st_row) + ":B" + str(st_row+rmax-5) + "").NumberFormat = "m/d/yyyy"

#     ws.Range(ws.Cells(st_row-4, st_col),ws.Cells(st_row-4, st_col)).Value = "DATE"
#     ws.Range(ws.Cells(st_row-4, st_col+1),ws.Cells(st_row-4, st_col+1)).Value = "CEK-KES BY PERSON"
  
    
#     isegment = int((cmax-1)/8) + 1
#     # if row['STORE_CODE'] == 'VTR' :
#       # print("here", cmax-1, isegment)

#     for i in range(1,isegment,1):
#       salpers = ws.Range(ws.Cells(st_row-1, st_col + ((i-1)*8) + 1), ws.Cells(st_row-1,st_col + ((i-1)*8) + 1)).Value
#       salpers = salpers.replace("SALES_VALUE_","").replace("_"," ")
#       # if row['STORE_CODE'] == 'VTR' :
#       #   print("ini", salpers)
#       # naming sales person
#       ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-3,st_col + ((i-1)*8) + 1)).Value = salpers
#       ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-3,st_col + ((i-1)*8) + 8)).MergeCells = True
 
#       ws.Range(ws.Cells(st_row-2, st_col + ((i-1)*8) + 1),ws.Cells(st_row-2, st_col + ((i-1)*8) + 1)).Value = "TOTAL"
#       ws.Range(ws.Cells(st_row-2, st_col + ((i-1)*8) + 3),ws.Cells(st_row-2, st_col + ((i-1)*8) + 3)).Value = "Gula Darah"
#       ws.Range(ws.Cells(st_row-2, st_col + ((i-1)*8) + 5),ws.Cells(st_row-2, st_col + ((i-1)*8) + 5)).Value = "Kolesterol"
#       ws.Range(ws.Cells(st_row-2, st_col + ((i-1)*8) + 7),ws.Cells(st_row-2, st_col + ((i-1)*8) + 7)).Value = "Asam Urat"

#       match i: 
#         case 1 :
#           ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-2,st_col + st_col + ((i-1)*8) + 5)).Interior.Color = rgbtohex((0,112,192)) #blue light
#         case 2 :
#           ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-2,st_col + st_col + ((i-1)*8) + 5)).Interior.Color = rgbtohex((112,48,160)) #purple
#         case 3 :
#           ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-2,st_col + st_col + ((i-1)*8) + 5)).Interior.Color = rgbtohex((75,172,198)) #green bean
#         case 4 :
#           ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-2,st_col + st_col + ((i-1)*8) + 5)).Interior.Color = rgbtohex((218,150,148)) #dark pink
#         case 5 :
#           ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-2,st_col + st_col + ((i-1)*8) + 5)).Interior.Color = rgbtohex((0,176,80)) #green
#         case 6 :
#           ws.Range(ws.Cells(st_row-3,st_col + ((i-1)*8) + 1), ws.Cells(st_row-2,st_col + st_col + ((i-1)*8) + 5)).Interior.Color = rgbtohex((247,150,70)) #orange  

#     for s in range (1,cmax,2):
#       ws.Range(ws.Cells(st_row-1, st_col + s),ws.Cells(st_row-1, st_col + s)).Value = "Sales"      
#       ws.Range(ws.Cells(st_row-2, st_col + s),ws.Cells(st_row-2, st_col + s + 1)).MergeCells = True

#     for s in range (2,cmax,2):
#       ws.Range(ws.Cells(st_row-1, st_col + s),ws.Cells(st_row-1, st_col + s)).Value = "Trx"

#     ws.Range(ws.Cells(st_row-4, st_col + 1),ws.Cells(st_row-4, st_col + cmax-1)).MergeCells = True
#     ws.Range(ws.Cells(st_row-4, st_col),ws.Cells(st_row-2, st_col)).MergeCells = True
#     ws.Range(ws.Cells(st_row-4, st_col),ws.Cells(st_row-2, st_col)).VerticalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-4, st_col),ws.Cells(st_row-2, st_col)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row, st_col+1),ws.Cells(st_row+rmax, st_col+cmax-1)).NumberFormat = "#,#0"
    
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Font.Name = "Calibri"
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Font.FontStyle = "Bold"
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Font.Size = 11
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Font.Color = rgbtohex((255,255,255)) # putih
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-4,st_col+cmax-1)).Interior.Color = rgbtohex((0,20,60)) #dark
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Interior.Color = rgbtohex((192,0,0)) #red
    
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-1,st_col+cmax-1)).HorizontalAlignment = c.xlCenter
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row-1,st_col+cmax-1)).VerticalAlignment = c.xlCenter

#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Borders.Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-4,st_col), ws.Cells(st_row-1,st_col+cmax-1)).Borders.Color = 0


#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders.LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders.Weight = c.xlHairline
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders.Color = 0

#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlInsideVertical).Weight = c.xlThin
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlInsideVertical).Color = 0

#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeTop).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeTop).Color = 0
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeBottom).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeBottom).Color = 0
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeLeft).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeLeft).Color = 0
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Borders(c.xlEdgeRight).Color = 0
    
#     ws.Range(ws.Cells(st_row-1,st_col), ws.Cells(st_row-1,st_col)).ClearContents()

#     # Grand Total
#     ws.Range(ws.Cells(st_row+rmax-4,st_col), ws.Cells(st_row+rmax-4,st_col+cmax-1)).Interior.Color = rgbtohex((196, 189, 151)) #broken white
#     # formula
#     for i in range(st_col+1,cmax+st_col,1):
#       ws.Range(ws.Cells(st_row+rmax-4,i), ws.Cells(st_row+rmax-4,i)).FormulaR1C1 = "=SUM(R[-" + str(rmax-4) + "]C:R[-1]C)"

#     # date double border
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row + rmax-4, st_col)).Borders(c.xlEdgeRight).LineStyle = c.xlDouble
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row + rmax-4, st_col)).Borders(c.xlEdgeRight).Weight = c.xlMedium
#     ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row + rmax-4, st_col)).Borders(c.xlEdgeRight).Color = 0
    
#     for i in range(1,isegment,1):
#       ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row + rmax-4, st_col + ((i-1)*8) )).Borders(c.xlEdgeRight).LineStyle = c.xlDouble
#       ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row + rmax-4, st_col + ((i-1)*8) )).Borders(c.xlEdgeRight).Weight = c.xlMedium
#       ws.Range(ws.Cells(st_row-3,st_col), ws.Cells(st_row + rmax-4, st_col + ((i-1)*8) )).Borders(c.xlEdgeRight).Color = 0


#     # if row['STORE_CODE'] == 'VTR' :
#     #   print ('VTR ok')  

    
#     ws.Range(ws.Cells(1,1), ws.Cells(1,1)).Copy()
#     ws.Range(ws.Cells(1,1), ws.Cells(1,1)).PasteSpecial(Paste=c.xlPasteValues)
#     ws.Range("A1").Select 
#     ws.Columns("Q:Q").ColumnWidth = 11
#     ws.Columns("R:R").ColumnWidth = 12
#     ws.Columns("J:J").ColumnWidth = 11
#     ws.Columns("U:U").ColumnWidth = 8

# END 
    wb.SaveAs(resultPath)



# ========================= SHEET SUMMARY MTD =========== 
ws = wb.Worksheets.Add()
ws.Name = "SUMMARY (MTD)"
xl.ActiveWindow.DisplayGridlines = False
xl.ActiveWindow.Zoom = 90
xl.DisplayAlerts = False

ws.Range("B1").Value = "SUMMARY MTD"  

q = "CALL SP_RPT_DailySalesMTD(14, null, null)"
dfw = pd.read_sql(q, cn)
dfw = dfw.fillna('')
    
cmax = len(dfw.columns)
rmax = len(dfw.index) + 4

# Paste recordset
st_row = 6
st_col = 1
ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
          ws.Cells(st_row + len(dfw.index) - 1,
                  st_col + len(dfw.columns) - 1) # No -1 for the index
          ).Value = dfw.to_records(index=False)

st_row = 27
st_col = 1
ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
          ws.Cells(st_row + len(dfw.index) - 1,
                  st_col + len(dfw.columns) - 1) # No -1 for the index
          ).Value = dfw.to_records(index=False)

# table border 1
ws.Range("B3:AC5").Borders(c.xlInsideHorizontal).LineStyle = c.xlContinuous
ws.Range("B3:AC5").Borders(c.xlInsideHorizontal).Weight = c.xlThin
ws.Range("B3:AC5").Borders(c.xlInsideHorizontal).Color = 0
ws.Range("B3:AC5").Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
ws.Range("B3:AC5").Borders(c.xlInsideVertical).Weight = c.xlThin
ws.Range("B3:AC5").Borders(c.xlInsideVertical).Color = 0
ws.Range("B3:AC5").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("B3:AC5").Borders(c.xlEdgeTop).Weight = c.xlThin
ws.Range("B3:AC5").Borders(c.xlEdgeTop).Color = 0
ws.Range("B3:AC5").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("B3:AC5").Borders(c.xlEdgeBottom).Weight = c.xlThin
ws.Range("B3:AC5").Borders(c.xlEdgeBottom).Color = 0
ws.Range("B3:AC5").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("B3:AC5").Borders(c.xlEdgeLeft).Weight = c.xlThin
ws.Range("B3:AC5").Borders(c.xlEdgeLeft).Color = 0
ws.Range("B3:AC5").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("B3:AC5").Borders(c.xlEdgeRight).Weight = c.xlThin
ws.Range("B3:AC5").Borders(c.xlEdgeRight).Color = 0
# Header Format
ws.Range("B3:AC5").Font.Name = "Calibri"
ws.Range("B3:AC5").Font.FontStyle = "Bold"
ws.Range("B3:AC5").Font.Size = 11
ws.Range("B3:AC5").HorizontalAlignment = c.xlCenter
ws.Range("B3:AC5").VerticalAlignment = c.xlCenter

ws.Range("B3:AC21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("B3:AC21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("B3:AC21").Borders(c.xlEdgeTop).Color = 0
ws.Range("B3:AC21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("B3:AC21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("B3:AC21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("B3:AC21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("B3:AC21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("B3:AC21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("B3:AC21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("B3:AC21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("B3:AC21").Borders(c.xlEdgeRight).Color = 0

ws.Range("B6:AC21").Borders(c.xlInsideHorizontal).LineStyle = c.xlContinuous
ws.Range("B6:AC21").Borders(c.xlInsideHorizontal).Weight = c.xlHairline
ws.Range("B6:AC21").Borders(c.xlInsideHorizontal).Color = 0
ws.Range("B6:AC21").Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
ws.Range("B6:AC21").Borders(c.xlInsideVertical).Weight = c.xlThin
ws.Range("B6:AC21").Borders(c.xlInsideVertical).Color = 0

ws.Range("C3").Value = "TOTAL DAILY SALES"
ws.Range("C3:K3").MergeCells = True
ws.Range("B3").Value = "STORE"
ws.Range("B3:B4").MergeCells = True
ws.Range("B5").Value = "TOTAL"
ws.Range("C4").Value = "Actual"
ws.Range("D4").Value = "Target"
ws.Range("E4").Value = "%Ach."
ws.Range("F4").Value = "Trx"
ws.Range("G4").Value = "Last Month Trx"
ws.Range("H4").Value = "vs LM"
ws.Range("I4").Value = "ABV"
ws.Range("J4").Value = "Last Month ABV"
ws.Range("K4").Value = "vs LM"

ws.Range("C3:K21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("C3:K21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("C3:K21").Borders(c.xlEdgeTop).Color = 0
ws.Range("C3:K21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("C3:K21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("C3:K21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("C3:K21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("C3:K21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("C3:K21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("C3:K21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("C3:K21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("C3:K21").Borders(c.xlEdgeRight).Color = 0

ws.Range("L3").Value = "HALODOC"
ws.Range("L3:N3").MergeCells = True
ws.Range("L4").Value = "Actual"
ws.Range("M4").Value = "Target"
ws.Range("N4").Value = "%Ach."

ws.Range("L3:N21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("L3:N21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("L3:N21").Borders(c.xlEdgeTop).Color = 0
ws.Range("L3:N21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("L3:N21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("L3:N21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("L3:N21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("L3:N21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("L3:N21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("L3:N21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("L3:N21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("L3:N21").Borders(c.xlEdgeRight).Color = 0

ws.Range("O3").Value = "CEK-KES"
ws.Range("O3:Q3").MergeCells = True
ws.Range("O4").Value = "Actual"
ws.Range("P4").Value = "Target"
ws.Range("Q4").Value = "%Ach."

ws.Range("O3:Q21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("O3:Q21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("O3:Q21").Borders(c.xlEdgeTop).Color = 0
ws.Range("O3:Q21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("O3:Q21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("O3:Q21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("O3:Q21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("O3:Q21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("O3:Q21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("O3:Q21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("O3:Q21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("O3:Q21").Borders(c.xlEdgeRight).Color = 0

ws.Range("R3").Value = "PWP"
ws.Range("R3:T3").MergeCells = True
ws.Range("R4").Value = "Actual"
ws.Range("S4").Value = "Target"
ws.Range("T4").Value = "%Ach."

ws.Range("R3:T21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("R3:T21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("R3:T21").Borders(c.xlEdgeTop).Color = 0
ws.Range("R3:T21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("R3:T21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("R3:T21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("R3:T21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("R3:T21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("R3:T21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("R3:T21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("R3:T21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("R3:T21").Borders(c.xlEdgeRight).Color = 0

ws.Range("U3").Value = "THEMATIC"
ws.Range("U3:W3").MergeCells = True
ws.Range("U4").Value = "Actual"
ws.Range("V4").Value = "Target"
ws.Range("W4").Value = "%Ach."

ws.Range("U3:W21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("U3:W21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("U3:W21").Borders(c.xlEdgeTop).Color = 0
ws.Range("U3:W21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("U3:W21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("U3:W21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("U3:W21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("U3:W21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("U3:W21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("U3:W21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("U3:W21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("U3:W21").Borders(c.xlEdgeRight).Color = 0


ws.Range("X3").Value = "BEST CHOICE"
ws.Range("X3:Z3").MergeCells = True
ws.Range("X4").Value = "Actual"
ws.Range("Y4").Value = "Target"
ws.Range("Z4").Value = "%Ach."

ws.Range("X3:Z21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("X3:Z21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("X3:Z21").Borders(c.xlEdgeTop).Color = 0
ws.Range("X3:Z21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("X3:Z21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("X3:Z21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("X3:Z21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("X3:Z21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("X3:Z21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("X3:Z21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("X3:Z21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("X3:Z21").Borders(c.xlEdgeRight).Color = 0

ws.Range("AA3").Value = "PHARMA"
ws.Range("AA3:AC3").MergeCells = True
ws.Range("AA4").Value = "Actual"
ws.Range("AB4").Value = "Target"
ws.Range("AC4").Value = "%Ach."

ws.Range("AA3:AC21").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("AA3:AC21").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("AA3:AC21").Borders(c.xlEdgeTop).Color = 0
ws.Range("AA3:AC21").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("AA3:AC21").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("AA3:AC21").Borders(c.xlEdgeBottom).Color = 0
ws.Range("AA3:AC21").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("AA3:AC21").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("AA3:AC21").Borders(c.xlEdgeLeft).Color = 0
ws.Range("AA3:AC21").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("AA3:AC21").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("AA3:AC21").Borders(c.xlEdgeRight).Color = 0

# table border 2
ws.Range("B24:N26").Borders(c.xlInsideHorizontal).LineStyle = c.xlContinuous
ws.Range("B24:N26").Borders(c.xlInsideHorizontal).Weight = c.xlThin
ws.Range("B24:N26").Borders(c.xlInsideHorizontal).Color = 0
ws.Range("B24:N26").Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
ws.Range("B24:N26").Borders(c.xlInsideVertical).Weight = c.xlThin
ws.Range("B24:N26").Borders(c.xlInsideVertical).Color = 0
ws.Range("B24:N26").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("B24:N26").Borders(c.xlEdgeTop).Weight = c.xlThin
ws.Range("B24:N26").Borders(c.xlEdgeTop).Color = 0
ws.Range("B24:N26").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("B24:N26").Borders(c.xlEdgeBottom).Weight = c.xlThin
ws.Range("B24:N26").Borders(c.xlEdgeBottom).Color = 0
ws.Range("B24:N26").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("B24:N26").Borders(c.xlEdgeLeft).Weight = c.xlThin
ws.Range("B24:N26").Borders(c.xlEdgeLeft).Color = 0
ws.Range("B24:N26").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("B24:N26").Borders(c.xlEdgeRight).Weight = c.xlThin
ws.Range("B24:N26").Borders(c.xlEdgeRight).Color = 0
# Header Format
ws.Range("B24:N26").Font.Name = "Calibri"
ws.Range("B24:N26").Font.FontStyle = "Bold"
ws.Range("B24:N26").Font.Size = 11
ws.Range("B24:N26").HorizontalAlignment = c.xlCenter
ws.Range("B24:N26").VerticalAlignment = c.xlCenter

ws.Range("B24:N42").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeTop).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeBottom).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeLeft).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeRight).Color = 0

ws.Range("B27:N42").Borders(c.xlInsideHorizontal).LineStyle = c.xlContinuous
ws.Range("B27:N42").Borders(c.xlInsideHorizontal).Weight = c.xlHairline
ws.Range("B27:N42").Borders(c.xlInsideHorizontal).Color = 0

ws.Range("B24:N42").Borders(c.xlInsideVertical).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlInsideVertical).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlInsideVertical).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeTop).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeBottom).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeLeft).Color = 0
ws.Range("B24:N42").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("B24:N42").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("B24:N42").Borders(c.xlEdgeRight).Color = 0

ws.Range("B24").Value = "STORE"
ws.Range("B24:B25").MergeCells = True
ws.Range("B26").Value = "TOTAL"

ws.Range("C24").Value = "NO OF EXISTING MEMBER"
ws.Range("C24:E24").MergeCells = True
ws.Range("C25").Value = "Actual"
ws.Range("D25").Value = "Target"
ws.Range("E25").Value = "%Ach."

ws.Range("C24:E42").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("C24:E42").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("C24:E42").Borders(c.xlEdgeTop).Color = 0
ws.Range("C24:E42").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("C24:E42").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("C24:E42").Borders(c.xlEdgeBottom).Color = 0
ws.Range("C24:E42").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("C24:E42").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("C24:E42").Borders(c.xlEdgeLeft).Color = 0
ws.Range("C24:E42").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("C24:E42").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("C24:E42").Borders(c.xlEdgeRight).Color = 0

ws.Range("F24").Value = "TRANSACTION EXISTING MEMBER"
ws.Range("F24:H24").MergeCells = True
ws.Range("F25").Value = "Actual"
ws.Range("G25").Value = "Target"
ws.Range("H25").Value = "%Ach."

ws.Range("F24:H42").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("F24:H42").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("F24:H42").Borders(c.xlEdgeTop).Color = 0
ws.Range("F24:H42").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("F24:H42").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("F24:H42").Borders(c.xlEdgeBottom).Color = 0
ws.Range("F24:H42").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("F24:H42").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("F24:H42").Borders(c.xlEdgeLeft).Color = 0
ws.Range("F24:H42").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("F24:H42").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("F24:H42").Borders(c.xlEdgeRight).Color = 0

ws.Range("I24").Value = "NO OF NEW MEMBER"
ws.Range("I24:K24").MergeCells = True
ws.Range("I25").Value = "Actual"
ws.Range("J25").Value = "Target"
ws.Range("K25").Value = "%Ach."

ws.Range("I24:K42").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("I24:K42").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("I24:K42").Borders(c.xlEdgeTop).Color = 0
ws.Range("I24:K42").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("I24:K42").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("I24:K42").Borders(c.xlEdgeBottom).Color = 0
ws.Range("I24:K42").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("I24:K42").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("I24:K42").Borders(c.xlEdgeLeft).Color = 0
ws.Range("I24:K42").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("I24:K42").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("I24:K42").Borders(c.xlEdgeRight).Color = 0

ws.Range("L24").Value = "TRANSACTION NEW+NON MEMBER"
ws.Range("L24:N24").MergeCells = True
ws.Range("L25").Value = "Actual"
ws.Range("M25").Value = "Target"
ws.Range("N25").Value = "%Ach."

ws.Range("L24:N42").Borders(c.xlEdgeTop).LineStyle = c.xlContinuous
ws.Range("L24:N42").Borders(c.xlEdgeTop).Weight = c.xlMedium
ws.Range("L24:N42").Borders(c.xlEdgeTop).Color = 0
ws.Range("L24:N42").Borders(c.xlEdgeBottom).LineStyle = c.xlContinuous
ws.Range("L24:N42").Borders(c.xlEdgeBottom).Weight = c.xlMedium
ws.Range("L24:N42").Borders(c.xlEdgeBottom).Color = 0
ws.Range("L24:N42").Borders(c.xlEdgeLeft).LineStyle = c.xlContinuous
ws.Range("L24:N42").Borders(c.xlEdgeLeft).Weight = c.xlMedium
ws.Range("L24:N42").Borders(c.xlEdgeLeft).Color = 0
ws.Range("L24:N42").Borders(c.xlEdgeRight).LineStyle = c.xlContinuous
ws.Range("L24:N42").Borders(c.xlEdgeRight).Weight = c.xlMedium
ws.Range("L24:N42").Borders(c.xlEdgeRight).Color = 0

ws.Columns("B:B").ColumnWidth = 40
ws.Columns("C:AC").ColumnWidth = 14

# finalize workbook

for sheet in wb.Sheets:
  xl.DisplayAlerts = False
  # print(sheet.Name)
  if sheet.Name == 'Sheet1':
    wb.Worksheets('Sheet1').Delete()
  elif sheet.Name == 'Sheet2':
    wb.Worksheets('Sheet2').Delete()
  elif sheet.Name == 'Sheet3':
    wb.Worksheets('Sheet3').Delete() 

wb.SaveAs(resultPath)
wb.Close()
xl.DisplayAlerts = True

# CLOSING...
cn.close()
xl.DisplayAlerts = True
xl.Quit()
# Output to ssis
# print(resultPath + ";" + dt_file)
print("done")

