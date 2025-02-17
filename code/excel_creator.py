import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment,Font

def func1(num):
    if num==1:
        return 'I'
    elif num==2:
        return 'II'
    elif num==3:
        return 'III'
    elif num==4:
        return 'IV'
    elif num==5:
        return 'V'
    elif num==6:
        return 'VI'
    elif num==7:
        return 'VII'
    else:
        return 'VIII'

def create(name):
    df = pd.read_excel(name)
    df.columns = df.columns.str.strip()
    students = df["BT ID"].unique()
    for student in students:
        student_df = df[df["BT ID"] == student].sort_values(by="Semester")
        output_filename = f"{student}_Grade_Card.xlsx"
        
        headers = []

        with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
            left_start_row = 4
            right_start_row = 4
            
            left_start_col = 1
            right_start_col = 11
            
            table_cols = ["Course Code","Course","Credit","Grade"]
            
            unique_semesters = sorted(student_df["Semester"].unique())

            for semester in unique_semesters:
                sem_df = student_df[student_df["Semester"] == semester]

                def func(sem_df,left_start_row,left_start_col):
                    sem_df["Course Code"].to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col, index=False)
                    sem_df["Course"].to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col+2, index=False)
                    sem_df["Credit"].to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col+6, index=False)
                    sem_df["Grade"].to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col+8, index=False)

                if int(semester) % 2 == 1:
                    headers.append((left_start_row + 1, left_start_col + 1,8))
                    header_text = f"SEM. {func1(semester)} (July-Nov {int('20'+student[2:4])+int(semester)//2})"  
                    header_df = pd.DataFrame({" ": [header_text]})
                    header_df.to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col, index=False, header=False)
                    
                    left_start_row += 1
                    func(sem_df,left_start_row,left_start_col)
                    for j in range(len(sem_df)+1):
                        headers.append((left_start_row+j+1, left_start_col + 1,1))
                        headers.append((left_start_row+j+1, left_start_col + 3,3))
                        headers.append((left_start_row+j+1, left_start_col + 7,1))

                    left_start_row += len(sem_df) + 2
                   
                else:
                    headers.append((right_start_row + 1, right_start_col + 1,8))
                    header_text = f"SEM. {func1(semester)} (Jan-May {int('20'+student[2:4])+int(semester)//2})"  
                    header_df = pd.DataFrame({" ": [header_text]})
                    header_df.to_excel(writer, sheet_name="Sheet1", startrow=right_start_row, startcol=right_start_col, index=False, header=False)
                    right_start_row += 1                    
                    func(sem_df,right_start_row,right_start_col)
                    for j in range(len(sem_df)+1):
                        headers.append((right_start_row+j+1, right_start_col + 1,1))
                        headers.append((right_start_row+j+1, right_start_col + 3,3))
                        headers.append((right_start_row+j+1, right_start_col + 7,1))
                    right_start_row += len(sem_df) + 2
        
        wb = load_workbook(output_filename)
        ws = wb.active

        ws.column_dimensions['A'].width = 1
        ws.column_dimensions['K'].width = 3
        # column_widths = {"A": 20, "B": 30, "C": 15}

        # for col, width in column_widths.items():
            # ws.column_dimensions[col].width = width
        # column_widths = {"C": 30, "H": 30, "B": 15,"G":15}

        bold_font = Font(bold=True)
        for row, col,k in headers:
            ws.cell(row=row, column=col).font = bold_font
            ws.merge_cells(start_row=row,end_row=row,start_column=col,end_column=col+k)

        # for col, width in column_widths.items():
        #     ws.column_dimensions[col].width = width

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center',wrap_text=True)
        
        wb.save(output_filename)

        print(f"Created file for {student}: {output_filename}")
        break
    print("Processing complete.")