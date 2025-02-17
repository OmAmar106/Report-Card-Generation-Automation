import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

def create(name):
    df = pd.read_excel(name)
    df.columns = df.columns.str.strip()
    students = df["BT ID"].unique()
    for student in students:
        student_df = df[df["BT ID"] == student].sort_values(by="Semester")
        output_filename = f"{student}_Grade_Card.xlsx"
        
        with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
            left_start_row = 4
            right_start_row = 4
            
            left_start_col = 1
            right_start_col = 6
            
            table_cols = ["Course Code","Course" ,"Credit","Grade"]
            
            unique_semesters = sorted(student_df["Semester"].unique())
            
            for semester in unique_semesters:
                sem_df = student_df[student_df["Semester"] == semester]
                header_text = f"Semester {semester}"
                
                header_df = pd.DataFrame({"" : [header_text]})
                
                if int(semester) % 2 == 1:
                    header_df.to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col, index=False, header=False)
                    left_start_row += 1
                    sem_df[table_cols].to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col, index=False)
                    
                    left_start_row += len(sem_df) + 2

                    if semester == 1:
                        blank_row = pd.DataFrame([["", "", "", ""]], columns=table_cols)
                        blank_row.to_excel(writer, sheet_name="Sheet1", startrow=left_start_row, startcol=left_start_col, index=False, header=False)
                        left_start_row += 1

                else:
                    header_df.to_excel(writer, sheet_name="Sheet1", startrow=right_start_row, startcol=right_start_col, index=False, header=False)
                    right_start_row += 1
                    
                    sem_df[table_cols].to_excel(writer, sheet_name="Sheet1", startrow=right_start_row, startcol=right_start_col, index=False)
                    
                    right_start_row += len(sem_df) + 2
        
        wb = load_workbook(output_filename)
        ws = wb.active
        
        # column_widths = {"A": 20, "B": 30, "C": 15}

        # for col, width in column_widths.items():
            # ws.column_dimensions[col].width = width
        column_widths = {"C": 25, "H": 25, "B": 15,"G":15}

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center',wrap_text=True)
        
        wb.save(output_filename)

        print(f"Created file for {student}: {output_filename}")
        break
    print("Processing complete.")