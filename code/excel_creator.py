
from openpyxl import Workbook 
from data_extraction import getalldata

def createexcel(btid,name,data):

    wb = Workbook()

    ws = wb.active


