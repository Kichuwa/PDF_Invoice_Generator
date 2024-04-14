from fpdf import FPDF
import pandas as pd
import glob

filepaths = glob.glob("invoices/*xlsx")

for path in filepaths:
    df = pd.read_excel(path, sheet_name="Sheet 1")

