import win32com.client
import os

def createpdf(input_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(input_path))
    ws = wb.Sheets(1)
    ws.PageSetup.PaperSize = 9
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.TopMargin = 150
    wb.ExportAsFixedFormat(0,os.path.abspath(input_path[:-4]+'pdf'))
    wb.Close(False)
    excel.Quit()
