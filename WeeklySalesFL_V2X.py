# to run
# py WeeklySalesFL.py 20220530
from calendar import month
import pandas as pd
import pypyodbc as pdb
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
# workdir = "F:\\03_REPORT\\04_WEEKLY\\Weekly Sales Hoka Principal"
workdir = "D:\\Ismail_MAA\\apps"
dt_aod = datetime.strptime(str(sys.argv[1]), '%Y%m%d')
dt_st = dt_aod - timedelta(days = dt_aod.weekday())
dt_st_lw = dt_st - timedelta(weeks = 2) 
dt_ed_lw = dt_st_lw + timedelta(days = 6) 
dt_st_cw = dt_st - timedelta(weeks = 1) 
dt_ed_cw = dt_st_cw + timedelta(days = 6) 

dt_aod_lm = dt_aod - relativedelta(months = 1)
dt_cm = date(dt_aod_lm.year + dt_aod_lm.month  // 12, dt_aod_lm.month % 12 + 1, 1) - timedelta(1)
dt_cm = datetime.strftime(dt_cm, '%b %Y')

dt_aod_lm = dt_aod - relativedelta(months = 2)
dt_lm = date(dt_aod_lm.year + dt_aod_lm.month  // 12, dt_aod_lm.month % 12 + 1, 1) - timedelta(1)
dt_lm = datetime.strftime(dt_lm, '%b %Y')
dt_stk_m = dt_cm
# Format date last week
if datetime.strftime(dt_st_lw, '%b') == datetime.strftime(dt_ed_lw, '%b') and datetime.strftime(dt_st_lw, '%Y') == datetime.strftime(dt_ed_lw, '%Y'):
  dt_lw = datetime.strftime(dt_st_lw, '%d') + ' - ' + datetime.strftime(dt_ed_lw, '%d') + ' ' + datetime.strftime(dt_ed_lw, '%b %Y')
elif datetime.strftime(dt_st_lw, '%b') != datetime.strftime(dt_ed_lw, '%b') and datetime.strftime(dt_st_lw, '%Y') == datetime.strftime(dt_ed_lw, '%Y'):
  dt_lw = datetime.strftime(dt_st_lw, '%d %b') + ' - ' + datetime.strftime(dt_ed_lw, '%d %b') + ' ' + datetime.strftime(dt_ed_lw, '%Y')
else :
  dt_lw = datetime.strftime(dt_st_lw, '%d %b %Y') + ' - ' + datetime.strftime(dt_ed_lw, '%d %b %Y')

# Format date current week
if datetime.strftime(dt_st_cw, '%b') == datetime.strftime(dt_ed_cw, '%b') and datetime.strftime(dt_st_cw, '%Y') == datetime.strftime(dt_ed_cw, '%Y'):
  dt_cw = datetime.strftime(dt_st_cw, '%d') + ' - ' + datetime.strftime(dt_ed_cw, '%d') + ' ' + datetime.strftime(dt_ed_cw, '%b %Y')
elif datetime.strftime(dt_st_cw, '%b') != datetime.strftime(dt_ed_cw, '%b') and datetime.strftime(dt_st_cw, '%Y') == datetime.strftime(dt_ed_cw, '%Y'):
  dt_cw = datetime.strftime(dt_st_cw, '%d %b') + ' - ' + datetime.strftime(dt_ed_cw, '%d %b') + ' ' + datetime.strftime(dt_ed_cw, '%Y')
else :
  dt_cw = datetime.strftime(dt_st_cw, '%d %b %Y') + ' - ' + datetime.strftime(dt_ed_cw, '%d %b %Y')

dt_stk = datetime.strftime(dt_ed_cw, '%d %b %Y')
dt_file = datetime.strftime(dt_st, '%d %b %Y')
# print(dt_cm, dt_lm)

# xl = mycom.Dispatch('Excel.Application')
xl = mycom.gencache.EnsureDispatch('Excel.Application')

resultPath = os.path.join(workdir,'WeeklyFootLocker_' + dt_file.replace(' ', '') +'.xlsx')
# print(resultPath)
xl.Visible = False

##########################============ WEEKLY REPORT
wb = xl.Workbooks.Add()
ws = wb.Worksheets('Sheet1')
cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
q = "EXEC SP_WEEKLY_MONTHLY_FL '6','', '" + str(sys.argv[1]) + "'"
dfstweek = pd.read_sql(q, cn)
dfstweek = dfstweek.reset_index()  # make sure indexes pair with number of rows
for index, row in dfstweek.iterrows():
    # print(row['site'], row['sitename'])
    ws = wb.Worksheets.Add()
    if row['site'] == "":
      ws.Name = row['sitename']
    else:  
      ws.Name = row['site']
    ws.Tab.ColorIndex = 3
    xl.ActiveWindow.DisplayGridlines = False
    xl.DisplayAlerts = False

    # cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
    q = "EXEC SP_WEEKLY_MONTHLY_FL '3','" + row['site'] + "','" + str(sys.argv[1]) + "'"
    dfw = pd.read_sql(q, cn)
    # Paste header
    st_row = 1
    st_col = 1
    for col in dfw.columns:
      ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
      st_col = st_col + 1

    # Paste recordset
    st_row = 5
    st_col = 1
    ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
             ws.Cells(st_row + len(dfw.index) - 1,
                      st_col + len(dfw.columns) - 1) # No -1 for the index
             ).Value = dfw.to_records(index=False)
    ws.Range("A1:AE1").ClearContents()
              
    cmax = len(dfw.columns)
    rmax = len(dfw.index) + 4

    # Header Format
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Name = "Calibri"
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.FontStyle = "Bold"
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Size = 11
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Color = rgbtohex((255,255,255))
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Interior.Color = rgbtohex((31,73,125))
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).HorizontalAlignment = c.xlCenter
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).VerticalAlignment = c.xlCenter
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.Color = 0

    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.Weight = c.xlThin
    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.Color = 0

    ws.Rows(1).RowHeight = 44.25
    ws.Rows(2).RowHeight = 17.25
    ws.Rows(3).RowHeight = 51
    ws.Rows(4).RowHeight = 17.25

    ws.Range("A2:B4").MergeCells = True
    ws.Range("A2").Value = "DEPARTEMENT" 
    if row['site'] == '':
      ws.Range("A1").Value = "Weekly Sales By Departement"
    else:     
      ws.Range("A1").Value = row['sitename'] + " - Weekly Sales By Departement"   
    ws.Range("A1").Font.Name = "Calibri"
    ws.Range("A1").Font.FontStyle = "Bold"
    ws.Range("A1").Font.Size = 18

    ws.Range("C2:Q2").MergeCells = True
    ws.Range("C2").Value = "QTY" 
    ws.Range("R2:AE2").MergeCells = True
    ws.Range("R2").Value = "Sales Amount - In Mio"  
    
    ws.Range("C3:J3").MergeCells = True
    ws.Range("C3").Value = "Sales Current Week \n" + dt_cw 
    ws.Range("R3:Y3").MergeCells = True
    ws.Range("R3").Value = "Sales Current Week \n" + dt_cw  

    ws.Range("C4").Value = "Mon"  
    ws.Range("D4").Value = "Tue"  
    ws.Range("E4").Value = "Wed"  
    ws.Range("F4").Value = "Thu"  
    ws.Range("G4").Value = "Fri"  
    ws.Range("H4").Value = "Sat"  
    ws.Range("I4").Value = "Sun"  
    ws.Range("J4").Value = "Total"

    ws.Range("R4").Value = "Mon"  
    ws.Range("S4").Value = "Tue"  
    ws.Range("T4").Value = "Wed"  
    ws.Range("U4").Value = "Thu"  
    ws.Range("V4").Value = "Fri"  
    ws.Range("W4").Value = "Sat"  
    ws.Range("X4").Value = "Sun"  
    ws.Range("Y4").Value = "Total"  
    
    ws.Range("K3:K4").MergeCells = True
    ws.Range("L3:L4").MergeCells = True
    ws.Range("M3:M4").MergeCells = True
    ws.Range("N3:N4").MergeCells = True
    ws.Range("O3:O4").MergeCells = True
    ws.Range("P3:P4").MergeCells = True
    ws.Range("Q3:Q4").MergeCells = True

    ws.Range("Z3:Z4").MergeCells = True
    ws.Range("AA3:AA4").MergeCells = True
    ws.Range("AB3:AB4").MergeCells = True
    ws.Range("AC3:AC4").MergeCells = True
    ws.Range("AD3:AD4").MergeCells = True
    ws.Range("AE3:AE4").MergeCells = True

    ws.Range("K3").Value = "Sales Last Week \n" + dt_lw  
    ws.Range("L3").Value = "Growth \n vs \n Last Week"  
    ws.Range("M3").Value = "MTD"  
    ws.Range("N3").Value = "MTD %Cont."  
    ws.Range("O3").Value = "Current Stock \n As Of \n" + dt_stk  
    ws.Range("P3").Value = "Stock Contribute"  
    ws.Range("Q3").Value = "Weekly \n Sellthrough"  
    
    ws.Range("Z3").Value = "Sales Last Week \n" + dt_lw  
    ws.Range("AA3").Value = "Growth \n vs \n Last Week"  
    ws.Range("AB3").Value = "MTD"  
    ws.Range("AC3").Value = "MTD %Cont."  
    ws.Range("AD3").Value = "Sales Last Month \n MTD"  
    ws.Range("AE3").Value = "Growth \n vs Last Month \n MTD"  

    #228, 223 236
    ws.Range(ws.Cells(5,3), ws.Cells(rmax,cmax)).Interior.Color = rgbtohex((228, 223, 236))
    ws.Range(ws.Cells(5,10), ws.Cells(rmax,10)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,14), ws.Cells(rmax,14)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,16), ws.Cells(rmax,16)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,17), ws.Cells(rmax,17)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,25), ws.Cells(rmax,25)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,27), ws.Cells(rmax,27)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,29), ws.Cells(rmax,29)).Interior.Color = rgbtohex((255, 255, 255))
    ws.Range(ws.Cells(5,31), ws.Cells(rmax,31)).Interior.Color = rgbtohex((255, 255, 255))

    # TOTALS Format
    rowgrp = 4
    rowdiv = 4
    rowgrand = 4
    for i in range(5,rmax+1,1):
      if ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Grp":
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
        ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value.upper()
        ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((197, 217, 241))
        if rowgrp + 1 != i:
          ws.Rows(str(rowgrp + 1) + ":" + str(i-1)).Group()
          # print(row['site'], str(rowgrp + 1) + ":" + str(i-1))
          # print(row['site'], str(rowdiv))
          if rowdiv < rowgrp + 1:
            rowgrp = i
          else:
            rowgrp = rowdiv

      elif ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Div" and ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value != "Grand Total":  
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
        ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = "Total " + ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value.upper()
        ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((141, 180, 226))    
        if rowdiv + 1 != i:
          ws.Rows(str(rowdiv + 1) + ":" + str(i-1)).Group()
          # print(row['site'], str(rowdiv + 1) + ":" + str(i-1))
          # print(row['site'], str(rowgrand))

          if rowgrand < rowdiv + 1:
            rowdiv = i 
            rowgrp = rowdiv
          else:
            rowdiv = rowgrand
        
      elif ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Div" and ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value == "Grand Total":  
        ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value
        ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Borders(c.xlBottom).Weight = c.xlMedium   
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.Color = rgbtohex((255,255,255))
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((31,73,125))
        ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
        k = i
        if rowgrand + 1 != i:
          ws.Rows(str(rowgrand+1) + ":" + str(i-1)).Group()
    
    ws.Columns("A:AE").AutoFit()
    ws.Range("A2").EntireColumn.ColumnWidth = 7.57
    ws.Range("B2").EntireColumn.ColumnWidth = 43
    ws.Range("M2").EntireColumn.ColumnWidth = 8
    ws.Range("AB2").EntireColumn.ColumnWidth = 8
    ws.Columns("C:J").ColumnWidth = 8
    ws.Columns("R:Y").ColumnWidth = 8
    
    ws.Range(ws.Cells(5,18), ws.Cells(rmax,26)).NumberFormat = "#,#.0,,"
    ws.Range(ws.Cells(5,28), ws.Cells(rmax,28)).NumberFormat = "#,#.0,,"
    ws.Range(ws.Cells(5,30), ws.Cells(rmax,30)).NumberFormat = "#,#.0,,"
    ws.Range(ws.Cells(5,12), ws.Cells(rmax,12)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,14), ws.Cells(rmax,14)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,16), ws.Cells(rmax,16)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,17), ws.Cells(rmax,17)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,27), ws.Cells(rmax,27)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,29), ws.Cells(rmax,29)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,31), ws.Cells(rmax,31)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,3), ws.Cells(rmax,11)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(5,13), ws.Cells(rmax,13)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(5,15), ws.Cells(rmax,15)).NumberFormat = "#,##0"
    

    # Formula
    for i in range(5,rmax+1,1): 
      ws.Range(ws.Cells(i,12), ws.Cells(rmax,12)).Formula = "=IFERROR((J" + str(i) + "-K" + str(i) + ")/K" + str(i) + ",\" \")"
      ws.Range(ws.Cells(i,14), ws.Cells(rmax,14)).Formula = "=IFERROR((M" + str(i) + "/$M$" + str(k) + "),\" \")"
      ws.Range(ws.Cells(i,16), ws.Cells(rmax,16)).Formula = "=IFERROR((O" + str(i) + "/$O$" + str(k) + "),\" \")"
      ws.Range(ws.Cells(i,17), ws.Cells(rmax,17)).Formula = "=IFERROR((J" + str(i) + "/O" + str(k) + "),\" \")"
      ws.Range(ws.Cells(i,27), ws.Cells(rmax,27)).Formula = "=IFERROR(((Y" + str(i) + "-Z" + str(i) + ")/Z" + str(i) + "),\" \")"
      ws.Range(ws.Cells(i,29), ws.Cells(rmax,29)).Formula = "=IFERROR((AB" + str(i) + "/$AB$" + str(k) + "),\" \")"
      ws.Range(ws.Cells(i,31), ws.Cells(rmax,31)).Formula = "=IFERROR(((AB" + str(i) + "-AD" + str(i) + ")/AD" + str(i) + "),\" \")"

    ws.Columns("R:X").Group()
    ws.Columns("C:I").Group()
    ws.Range("C5").Select()
    xl.ActiveWindow.FreezePanes = True
    wb.SaveAs(resultPath)

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
# wb.Close()

##########################============ MONTHLY & YTD REPORT ===========================##########################################
# MONTHLY

# resultPathM = os.path.join(workdir,'Monthly_YTD_FootLocker_' + dt_cm.replace(' ', '') +'.xlsx')
# cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
q = "EXEC SP_WEEKLY_MONTHLY_FL '7','', '" + str(sys.argv[1]) + "'"
dfstmnth = pd.read_sql(q, cn)
dfstmnth = dfstmnth.reset_index()  # make sure indexes pair with number of rows
for index, row in dfstmnth.iterrows():
    # print(row['site'], row['sitename'])
    ws = wb.Worksheets.Add()
    if row['site'] == "":
      ws.Name = row['sitename']
    else:  
      ws.Name = 'MTD_' + row['site']
    ws.Tab.ColorIndex = 6
    xl.ActiveWindow.DisplayGridlines = False
    xl.DisplayAlerts = False

# wb = xl.Workbooks.Add()
# ws = wb.Worksheets('Sheet1')
# ws.Name = "Monthly Sales By Dept"
# ws.Tab.ColorIndex = 3
# xl.ActiveWindow.DisplayGridlines = False
# xl.DisplayAlerts = False

# cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
    q = "EXEC SP_WEEKLY_MONTHLY_FL '4','" + row['site'] + "','" + str(sys.argv[1]) + "'"
    dfw = pd.read_sql(q, cn)
    # Paste header
    st_row = 1
    st_col = 1
    for col in dfw.columns:
        ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
        st_col = st_col + 1

    # Paste recordset
    st_row = 5
    st_col = 1
    ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
            ws.Cells(st_row + len(dfw.index) - 1,
                    st_col + len(dfw.columns) - 1) # No -1 for the index
            ).Value = dfw.to_records(index=False)
    ws.Range("A1:M1").ClearContents()
            
    cmax = len(dfw.columns)
    rmax = len(dfw.index) + 4

    # Header Format
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Name = "Calibri"
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.FontStyle = "Bold"
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Size = 11
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Color = rgbtohex((255,255,255))
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Interior.Color = rgbtohex((31,73,125)) # biru tua
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).HorizontalAlignment = c.xlCenter
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).VerticalAlignment = c.xlCenter
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.Color = 0

    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.Weight = c.xlThin
    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.Color = 0

    ws.Rows(1).RowHeight = 44.25
    ws.Rows(2).RowHeight = 17.25
    ws.Rows(3).RowHeight = 51
    ws.Rows(4).RowHeight = 17.25

    ws.Range("A2:B4").MergeCells = True
    ws.Range("A2").Value = "DEPARTEMENT" 
    # ws.Range("A1").Value = "Monthly Sales By Departement"
    if row['site'] == '':
        ws.Range("A1").Value = "Monthly Sales By Departement"
    else:     
        ws.Range("A1").Value = row['sitename'] + " - Monthly Sales By Departement"   
    ws.Range("A1").Font.Name = "Calibri"
    ws.Range("A1").Font.FontStyle = "Bold"
    ws.Range("A1").Font.Size = 18

    ws.Range("C2:I2").MergeCells = True
    ws.Range("C2").Value = "QTY" 
    ws.Range("J2:M2").MergeCells = True
    ws.Range("J2").Value = "Sales Amount - In Mio"  

    ws.Range("C3:C4").MergeCells = True
    ws.Range("C3").Value = "Sales Current Month \n" + dt_cm 
    ws.Range("D3:D4").MergeCells = True
    ws.Range("D3").Value = "Sales Last Month \n" + dt_lm
    ws.Range("E3:E4").MergeCells = True  
    ws.Range("E3").Value = "Growth \n vs \n Last Month" 
    ws.Range("F3:F4").MergeCells = True   
    ws.Range("F3").Value = "MTD %Cont."  
    ws.Range("G3:G4").MergeCells = True   
    ws.Range("G3").Value = "Current Stock \n EOM \n" + dt_stk_m  
    ws.Range("H3:H4").MergeCells = True   
    ws.Range("H3").Value = "Stock Contribute"  
    ws.Range("I3:I4").MergeCells = True   
    ws.Range("I3").Value = "Monthly \n Sellthrough"  

    ws.Range("J3:J4").MergeCells = True   
    ws.Range("J3").Value = "Sales Current Month \n" + dt_cm  
    ws.Range("K3:K4").MergeCells = True   
    ws.Range("K3").Value = "Sales Last Month \n" + dt_lm 
    ws.Range("L3:L4").MergeCells = True   
    ws.Range("L3").Value = "Growth \n vs \n Last Month"   
    ws.Range("M3:M4").MergeCells = True   
    ws.Range("M3").Value = "MTD %Cont."  

    #228, 223 236 = ungu muda
    ws.Range(ws.Cells(5,3), ws.Cells(rmax,3)).Interior.Color = rgbtohex((228, 223, 236))
    ws.Range(ws.Cells(5,4), ws.Cells(rmax,4)).Interior.Color = rgbtohex((228, 223, 236))
    ws.Range(ws.Cells(5,7), ws.Cells(rmax,7)).Interior.Color = rgbtohex((228, 223, 236))
    ws.Range(ws.Cells(5,10), ws.Cells(rmax,10)).Interior.Color = rgbtohex((228, 223, 236))
    ws.Range(ws.Cells(5,11), ws.Cells(rmax,11)).Interior.Color = rgbtohex((228, 223, 236))

    # TOTALS Format
    rowgrp = 4
    rowdiv = 4
    rowgrand = 4
    for i in range(5,rmax+1,1):
        if ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Grp":
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
            ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value.upper()
            ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((197, 217, 241))
            if rowgrp + 1 != i:
                ws.Rows(str(rowgrp + 1) + ":" + str(i-1)).Group()
            if rowdiv < rowgrp + 1:
                rowgrp = i
            else:
                rowgrp = rowdiv

        elif ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Div" and ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value != "Grand Total":  
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
            ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = "Total " + ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value.upper()
            ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((141, 180, 226))    
            if rowdiv + 1 != i:
                ws.Rows(str(rowdiv + 1) + ":" + str(i-1)).Group()

            if rowgrand < rowdiv + 1:
                rowdiv = i 
                rowgrp = rowdiv
            else:
                rowdiv = rowgrand
        
        elif ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Div" and ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value == "Grand Total":  
            ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value
            ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Borders(c.xlBottom).Weight = c.xlMedium   
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.Color = rgbtohex((255,255,255))
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((31,73,125))
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
            k = i
            if rowgrand + 1 != i:
                ws.Rows(str(rowgrand+1) + ":" + str(i-1)).Group()

    ws.Columns("A:M").AutoFit()
    ws.Range("A2").EntireColumn.ColumnWidth = 7.57
    ws.Range("B2").EntireColumn.ColumnWidth = 43
    ws.Columns("C:M").ColumnWidth = 15.57

    # NUMBER FORMATS 
    ws.Range(ws.Cells(5,10), ws.Cells(rmax,11)).NumberFormat = "#,#.0,,"
    ws.Range(ws.Cells(5,5), ws.Cells(rmax,6)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,8), ws.Cells(rmax,9)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,12), ws.Cells(rmax,13)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(5,3), ws.Cells(rmax,4)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(5,7), ws.Cells(rmax,7)).NumberFormat = "#,##0"
    
    # Formula
    for i in range(5,rmax+1,1): 
        ws.Range(ws.Cells(i,5), ws.Cells(rmax,5)).Formula = "=IFERROR((C" + str(i) + "-D" + str(i) + ")/D" + str(i) + ",\" \")"
        ws.Range(ws.Cells(i,6), ws.Cells(rmax,6)).Formula = "=IFERROR((C" + str(i) + "/$C$" + str(k) + "),\" \")"
        ws.Range(ws.Cells(i,8), ws.Cells(rmax,8)).Formula = "=IFERROR((G" + str(i) + "/$G$" + str(k) + "),\" \")"
        ws.Range(ws.Cells(i,9), ws.Cells(rmax,9)).Formula = "=IFERROR((C" + str(i) + "/G" + str(i) + "),\" \")"
        ws.Range(ws.Cells(i,12), ws.Cells(rmax,12)).Formula = "=IFERROR(((J" + str(i) + "-K" + str(i) + ")/K" + str(i) + "),\" \")"
        ws.Range(ws.Cells(i,13), ws.Cells(rmax,13)).Formula = "=IFERROR((J" + str(i) + "/$J$" + str(k) + ")," ")"

    ws.Range("C5").Select()
    xl.ActiveWindow.FreezePanes = True
    wb.SaveAs(resultPath)

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
###========================= YEAR TO DATE=========================
dt_m1 = datetime.strftime(date(dt_aod.year, 1, 1), '%b-%Y')
dt_m2 = datetime.strftime(date(dt_aod.year, 2, 1), '%b-%Y')
dt_m3 = datetime.strftime(date(dt_aod.year, 3, 1), '%b-%Y')
dt_m4 = datetime.strftime(date(dt_aod.year, 4, 1), '%b-%Y')
dt_m5 = datetime.strftime(date(dt_aod.year, 5, 1), '%b-%Y')
dt_m6 = datetime.strftime(date(dt_aod.year, 6, 1), '%b-%Y')
dt_m7 = datetime.strftime(date(dt_aod.year, 7, 1), '%b-%Y')
dt_m8 = datetime.strftime(date(dt_aod.year, 8, 1), '%b-%Y')
dt_m9 = datetime.strftime(date(dt_aod.year, 9, 1), '%b-%Y')
dt_m10 = datetime.strftime(date(dt_aod.year, 10, 1), '%b-%Y')
dt_m11 = datetime.strftime(date(dt_aod.year, 11, 1), '%b-%Y')
dt_m12 = datetime.strftime(date(dt_aod.year, 12, 1), '%b-%Y')

yy = dt_aod.year
lyy = dt_aod.year - 1

ws = wb.Worksheets.Add()
ws.Name = "Sheet1"
ws.Tab.ColorIndex = 3
xl.ActiveWindow.DisplayGridlines = False

# cn = pdb.connect("DRIVER={SQL Server};Server=jkthomaasql;DATABASE=DMS_BI;Trusted_Connection=Yes")
#10b5f1
#ffff1d
q = "EXEC SP_WEEKLY_MONTHLY_FL '8','', '" + str(sys.argv[1]) + "'"
dfstyr = pd.read_sql(q, cn)
dfstyr = dfstyr.reset_index()  # make sure indexes pair with number of rows
for index, row in dfstyr.iterrows():
    # print(row['site'], row['sitename'])
    ws = wb.Worksheets.Add()
    if row['site'] == "":
      ws.Name = row['sitename']
    else:  
      ws.Name = 'YTD_' + row['site']
    ws.Tab.ColorIndex = 8
    xl.ActiveWindow.DisplayGridlines = False
    xl.DisplayAlerts = False

    q = "EXEC SP_WEEKLY_MONTHLY_FL '5','" + row['site'] + "','" + str(sys.argv[1]) + "'"
    dfw = pd.read_sql(q, cn)

    # Paste header
    st_row = 1
    st_col = 1
    for col in dfw.columns:
        ws.Range(ws.Cells(st_row, st_col),ws.Cells(st_row, st_col)).Value = col
        st_col = st_col + 1

    # Paste recordset
    st_row = 5
    st_col = 1
    ws.Range(ws.Cells(st_row, st_col),# Cell to start the "paste"
            ws.Cells(st_row + len(dfw.index) - 1,
                    st_col + len(dfw.columns) - 1) # No -1 for the index
            ).Value = dfw.to_records(index=False)
    ws.Range("A1:AF1").ClearContents()
            
    cmax = len(dfw.columns)
    rmax = len(dfw.index) + 4
    # print(cmax, rmax)
    # Header Format
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Name = "Calibri"
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.FontStyle = "Bold"
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Size = 11
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Font.Color = rgbtohex((255,255,255))
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Interior.Color = rgbtohex((31,73,125)) # biru tua
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).HorizontalAlignment = c.xlCenter
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).VerticalAlignment = c.xlCenter
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.Weight = c.xlMedium
    ws.Range(ws.Cells(2,1), ws.Cells(4,cmax)).Borders.Color = 0

    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.LineStyle = c.xlContinuous
    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.Weight = c.xlThin
    ws.Range(ws.Cells(5,1), ws.Cells(rmax,cmax)).Borders.Color = 0

    ws.Rows(1).RowHeight = 44.25
    ws.Rows(2).RowHeight = 17.25
    ws.Rows(3).RowHeight = 51
    ws.Rows(4).RowHeight = 17.25

    ws.Range("A2:B4").MergeCells = True
    ws.Range("A2").Value = "DEPARTEMENT" 
    # ws.Range("A1").Value = "Year To Date Sales By Departement"
    if row['site'] == '':
        ws.Range("A1").Value = "Year To Date Sales By Departement"
    else:     
        ws.Range("A1").Value = row['sitename'] + " - Year To Date Sales By Departement"   
    ws.Range("A1").Font.Name = "Calibri"
    ws.Range("A1").Font.FontStyle = "Bold"
    ws.Range("A1").Font.Size = 18

    ws.Range("C2:D2").MergeCells = True
    ws.Range("E2:F2").MergeCells = True
    ws.Range("G2:H2").MergeCells = True
    ws.Range("I2:J2").MergeCells = True
    ws.Range("K2:L2").MergeCells = True
    ws.Range("M2:N2").MergeCells = True
    ws.Range("O2:P2").MergeCells = True
    ws.Range("Q2:R2").MergeCells = True
    ws.Range("S2:T2").MergeCells = True
    ws.Range("U2:V2").MergeCells = True
    ws.Range("W2:X2").MergeCells = True
    ws.Range("Y2:Z2").MergeCells = True
    ws.Range("AA2:AB2").MergeCells = True
    ws.Range("AC2:AD2").MergeCells = True
    ws.Range("AE2:AF2").MergeCells = True
    ws.Range("C2").Value = dt_m1 
    ws.Range("E2").Value = dt_m2 
    ws.Range("G2").Value = dt_m3 
    ws.Range("I2").Value = dt_m4 
    ws.Range("K2").Value = dt_m5 
    ws.Range("M2").Value = dt_m6 
    ws.Range("O2").Value = dt_m7 
    ws.Range("Q2").Value = dt_m8 
    ws.Range("S2").Value = dt_m9 
    ws.Range("U2").Value = dt_m10 
    ws.Range("W2").Value = dt_m11 
    ws.Range("Y2").Value = dt_m12 
    ws.Range("AA2").Value = "Year To Date \n" + str(yy) 
    ws.Range("AC2").Value =  "Year To Date \n" + str(lyy)
    ws.Range("AE2").Value =  "Growth \n vs \n Last Year"

    for i in range(3,cmax + 1,1):
        ws.Range(ws.Cells(2,i), ws.Cells(3,i)).MergeCells = True
        if i % 2 == 0:
            ws.Range(ws.Cells(4,i), ws.Cells(4,i)).Value = "Qty"
        else:
            ws.Range(ws.Cells(4,i), ws.Cells(4,i)).Value = "Sales"

    # 228, 223 236 = ungu muda
    ws.Range(ws.Cells(5,3), ws.Cells(rmax,cmax-2)).Interior.Color = rgbtohex((228, 223, 236))

    # TOTALS Format
    rowgrp = 4
    rowdiv = 4
    rowgrand = 4
    for i in range(5,rmax+1,1):
        if ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Grp":
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
            ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value.upper()
            ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((197, 217, 241))
            if rowgrp + 1 != i:
                ws.Rows(str(rowgrp + 1) + ":" + str(i-1)).Group()
            if rowdiv < rowgrp + 1:
                rowgrp = i
            else:
                rowgrp = rowdiv

        elif ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Div" and ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value != "Grand Total":  
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
            ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = "Total " + ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value.upper()
            ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((141, 180, 226))    
            if rowdiv + 1 != i:
                ws.Rows(str(rowdiv + 1) + ":" + str(i-1)).Group()

            if rowgrand < rowdiv + 1:
                rowdiv = i 
                rowgrp = rowdiv
            else:
                rowdiv = rowgrand
            
        elif ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value == "Div" and ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value == "Grand Total":  
            ws.Range(ws.Cells(i,1), ws.Cells(i,1)).Value = ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value
            ws.Range(ws.Cells(i,2), ws.Cells(i,2)).Value = ""
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Borders(c.xlBottom).Weight = c.xlMedium   
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.Color = rgbtohex((255,255,255))
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Interior.Color = rgbtohex((31,73,125))
            ws.Range(ws.Cells(i,1), ws.Cells(i,cmax)).Font.FontStyle = "Bold"
            k = i
            if rowgrand + 1 != i:
                ws.Rows(str(rowgrand+1) + ":" + str(i-1)).Group()

        ws.Columns("A:AF").AutoFit()
        ws.Range("A2").EntireColumn.ColumnWidth = 7.57
        ws.Range("B2").EntireColumn.ColumnWidth = 43
        ws.Columns("C:AD").ColumnWidth = 8.43

        # NUMBER FORMATS 

    for i in range(3,cmax-2,1):
        ws.Range(ws.Cells(2,i), ws.Cells(3,i)).MergeCells = True
        if i % 2 == 0:
            ws.Range(ws.Cells(5,i), ws.Cells(rmax,i)).NumberFormat = "#,##0"
        else:  
            ws.Range(ws.Cells(5,i), ws.Cells(rmax,i)).NumberFormat = "#,#.0,,"
        
    ws.Range(ws.Cells(5,31), ws.Cells(rmax,32)).NumberFormat = "0.00%"
        
        # Formula
    for i in range(5,rmax+1,1): 
        ws.Range(ws.Cells(i,31), ws.Cells(rmax,31)).Formula = "=IFERROR(((AA" + str(i) + "-AC" + str(i) + ")/AC" + str(i) + "),\" \")"
        ws.Range(ws.Cells(i,32), ws.Cells(rmax,32)).Formula = "=IFERROR(((AB" + str(i) + "-AD" + str(i) + ")/AD" + str(i) + "),\" \")"


    ws.Range("C5").Select()
    xl.ActiveWindow.FreezePanes = True
    wb.SaveAs(resultPath)
    xl.DisplayAlerts = False
for sheet in wb.Sheets:
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
print(resultPath + ";" + dt_file)

