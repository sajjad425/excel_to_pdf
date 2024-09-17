#import library
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.units import inch

# Path to the Excel file and PDF file
excel_file_path = r'D:\Programming\Source Code\basic_info.xlsx'
pdf_file_path = r'D:\Programming\Source Code\output9.pdf'

# Load the Excel file
df = pd.read_excel(excel_file_path)

# Convert DataFrame to list of lists (data)
data = [df.columns.tolist()] + df.values.tolist()

# Create a PDF document
doc = SimpleDocTemplate(pdf_file_path, pagesize=letter)

# Create a table and apply styling
table = Table(data)
table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 0), (-1, -1), 10),
    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
]))

# Build the PDF document
doc.build([table])

print(f"PDF generated successfully at: {pdf_file_path}")
