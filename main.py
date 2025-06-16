import pandas as pd
import glob
import os
from fpdf import FPDF
from pathlib import Path

# Get the directory where the script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INVOICE_OUTPUT_DIR = os.path.join(SCRIPT_DIR, "pdf-invoices")

# Create output directory if it doesn't exist
os.makedirs(INVOICE_OUTPUT_DIR, exist_ok=True)

# Use os.path for reliable path handling
filepaths = glob.glob(os.path.join(SCRIPT_DIR, "invoices", "*.xlsx"))

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)

    filename = Path(filepath).stem
    invoice_number = filename.split('-')[0] if '-' in filename else filename

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(w=0, h=10, txt=f"Invoice nr. {invoice_number}", ln=True, align='L')

    pdf.output(os.path.join(INVOICE_OUTPUT_DIR, f"{filename}.pdf"))
