from data_extraction import getalldata
from excel_creator import createexcel

data,name = getalldata()

for i in data:
    createexcel(i,name[i],data[i])

