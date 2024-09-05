# Import the win32com module
from win32com import client 

# Open Microsoft Excel
excel = client.Dispatch("Excel.Application")

# Make Excel visible (optional)
excel.Visible = False

# Path to the Excel file you want to convert
excel_file_path = r'D:\Programming\Source Code\basic_info.xlsx'

# Path where the PDF will be saved
pdf_file_path = r'D:\Programming\Source Code\output.pdf'

# Open the Excel file
sheets = excel.Workbooks.Open(excel_file_path)

# Select the first worksheet (index 0)
work_sheets = sheets.Worksheets[0]

# Convert the worksheet to PDF
work_sheets.ExportAsFixedFormat(0, pdf_file_path)

# Close the Excel file
sheets.Close(False)

# Quit the Excel application
excel.Quit()

print(f"Conversion complete! The PDF has been saved at: {pdf_file_path}")