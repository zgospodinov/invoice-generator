import pandas as pd
import glob
import os
from fpdf import FPDF
from pathlib import Path
from datetime import datetime


# Get the directory where the script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INVOICE_OUTPUT_DIR = os.path.join(SCRIPT_DIR, "pdf-invoices")

# Create output directory if it doesn't exist
os.makedirs(INVOICE_OUTPUT_DIR, exist_ok=True)

# Get current date
current_date = datetime.now()

# Use os.path for reliable path handling
filepaths = glob.glob(os.path.join(SCRIPT_DIR, "invoices", "*.xlsx"))

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_number, invoice_date = filename.split('-') if '-' in filename else (filename, f"{current_date.strftime('%Y.%m.%d')}")

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(w=0, h=8, txt=f"Invoice nr. {invoice_number}", ln=True, align='L')
    pdf.cell(w=0, h=8, txt=f"Date: {invoice_date}", ln=True, align='L')
    
    # Add space after the date
    pdf.ln(10)  # Adds 10mm of vertical space

    # Read the Excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Debug print
    print(f"Columns found in Excel file: {list(df.columns)}")

    # Get columns from the dataframe
    columns = df.columns

    # Define column widths based on content
    col_widths = {}
    for column in columns:
        # Assign widths based on column type or name
        if 'name' in column.lower():
            col_widths[column] = 60  # Wider for name columns
        elif 'price' in column.lower() or 'amount' in column.lower():
            col_widths[column] = 35  # Medium for numeric columns
        else:
            col_widths[column] = 25  # Default width

    # Add table headers
    pdf.set_font("Arial", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.set_fill_color(240, 240, 240)
    
    for column, width in col_widths.items():
        header_text = column.replace('_', ' ').title()
        pdf.cell(w=width, h=8, txt=header_text, border=1, align='L', fill=True)
    pdf.ln()

    # Add table content
    pdf.set_font("Arial", size=10)
    for index, row in df.iterrows():
        pdf.set_text_color(80, 80, 80)

        for column, width in col_widths.items():
            value = str(row[column])
            align = 'R' if 'price' in column.lower() or 'amount' in column.lower() else 'L'
            pdf.cell(w=width, h=8, txt=value, border=1, align=align)
        pdf.ln()  # Move to next row

    # Calculate total
    total_sum = df['total_price'].sum() if 'total_price' in df.columns else df[df.columns[df.columns.str.contains('total', case=False)]].sum().iloc[0]
    
    # Add a small space before total row
    pdf.ln(5)
    
    # Add total row with bold font and gray background
    pdf.set_font("Arial", size=10, style='B')
    pdf.set_fill_color(240, 240, 240)
    
    # Add empty cells for all columns except the last two
    remaining_width = sum([width for col, width in col_widths.items() 
                         if not (col == 'total_price' or 'total' in col.lower())])
    pdf.cell(w=remaining_width, h=8, txt="Total:", border=1, align='R', fill=True)
    
    # Add total amount
    pdf.cell(w=list(col_widths.values())[-1], h=8, 
             txt=f"${total_sum:.2f}", border=1, align='R', fill=True)
    pdf.ln()

    # Add space after table
    pdf.ln(20)  # 20mm space

    # Add total due section with larger font
    pdf.set_font("Arial", size=16, style='B')
    pdf.set_text_color(50, 50, 50)
    pdf.cell(w=0, h=10, txt=f"Total Amount Due: ${total_sum:.2f}", ln=True, align='R')
    
    # Add space before company info
    pdf.ln(15)

    # Add company information
    pdf.set_font("Arial", size=14, style='B')
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=10, txt="TechCorp Solutions Ltd.", ln=True, align='L')
    pdf.image(os.path.join(SCRIPT_DIR, "pythonhow.png"), w=10, h=10)  # Adjust path and size as needed

    # Add company slogan or additional info in italic
    pdf.set_font("Arial", size=10, style='I')
    pdf.set_text_color(150, 150, 150)
    pdf.cell(w=0, h=8, txt="Innovating for tomorrow", ln=True, align='L')

    # Save the PDF
    pdf.output(os.path.join(INVOICE_OUTPUT_DIR, f"{filename}.pdf"))
