import pandas as pd
import glob
import os

# Get the directory where the script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Use os.path for reliable path handling
filepaths = glob.glob(os.path.join(SCRIPT_DIR, "invoices", "*.xlsx"))

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
