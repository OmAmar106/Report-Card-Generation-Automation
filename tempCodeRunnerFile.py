import os
import win32com.client

def convert_xlsx_to_pdf(input_xlsx, output_pdf_dir):
    # Open Excel application
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Run in the background

    try:
        # Open the Excel file
        workbook = excel.Workbooks.Open(os.path.abspath(input_xlsx))
        
        # Get the sheet names
        sheet_names = [sheet.Name for sheet in workbook.Sheets]
        
        for sheet_name in sheet_names:
            worksheet = workbook.Sheets(sheet_name)  # Select each sheet by name

            # **Adjust Page Setup**
            worksheet.PageSetup.Zoom = False  # Disable zoom
            worksheet.PageSetup.FitToPagesWide = 1  # Fit to one page width
            worksheet.PageSetup.FitToPagesTall = False  # Adjust height dynamically
            
            # Set orientation to landscape for better width handling
            worksheet.PageSetup.Orientation = 2  # 2 = Landscape, 1 = Portrait
            
            # Optional: Set Paper Size (e.g., A4 = 9, Letter = 1)
            worksheet.PageSetup.PaperSize = 9  # A4 size

            # Define the PDF file path for each sheet
            pdf_file = os.path.join(output_pdf_dir, f"{sheet_name}.pdf")

            # Export sheet as PDF
            worksheet.ExportAsFixedFormat(0, os.path.abspath(pdf_file))

            print(f"Successfully converted {sheet_name} to {pdf_file}")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # Close workbook without saving and quit Excel
        workbook.Close(SaveChanges=False)
        excel.Quit()

# Example usage
excel_file = 'C:/Users/asus/Desktop/sample.xlsx'  # Replace with your Excel file path
output_dir = 'C:/Users/asus/Desktop/pdf_outputs'  # Directory to save PDFs
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

convert_xlsx_to_pdf(excel_file, output_dir)
