import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add()
wb.SaveAs('D:\\ISMAIL_ERABW\\pyscript\\output\\add_a_workbook.xlsx')
excel.Application.Quit()