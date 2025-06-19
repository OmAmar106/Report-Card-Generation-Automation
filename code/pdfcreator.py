import win32com.client
import os

def createpdf(input_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(os.path.abspath(input_path))
    ws = wb.Sheets(1)
    
    for row in range(1, ws.UsedRange.Rows.Count + 1):  
        ws.Rows(row).RowHeight *= 1.2

    ws.PageSetup.PaperSize = 9
    ws.PageSetup.Zoom = 58
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.TopMargin = 20
    ws.PageSetup.LeftMargin = 20
    ws.PageSetup.RightMargin = 15
    ws.PageSetup.BottomMargin = 5



    wb.ExportAsFixedFormat(0, os.path.abspath(input_path[:-4] + 'pdf'))
    wb.Close(False)
    excel.Quit()