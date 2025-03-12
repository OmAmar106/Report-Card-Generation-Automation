from data_extraction import getalldata
from excel_creator import create
from openpyxl import Workbook
from openpyxl.styles import Alignment

# file_name = input("Enter File Name : ")
file_name = "ECE-2020-CUMULATIVE.xlsx"

data,name = getalldata(file_name)
def createexcel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Student Data"
    headers = ["BT ID","Student Name", "Semester", "Course", "Course Code", "Credit", "Grade"]
    ws.append(headers)
    for student, student_data in finaldata.items():
        full_name = name.get(student, "Unknown Name")
        for data_tuple in student_data:
            ws.append([student,full_name, data_tuple[0], data_tuple[1], data_tuple[2], data_tuple[3], data_tuple[4]])

    wb.save(filename)


filename = "student_data.xlsx"  
finaldata, name = getalldata(file_name)

createexcel()

create("student_data.xlsx","Electronics and Communication Engineering",name)

# create("student_data.xlsx",input("Enter Branch Name"))
