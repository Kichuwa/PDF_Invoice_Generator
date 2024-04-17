from fpdf import FPDF
import pandas as pd
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*xlsx")

for path in filepaths:
    df = pd.read_excel(path, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(path).stem
    invoice_number = filename.split("-")[0]
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice no. {invoice_number}")
    pdf.ln()

    invoice_date = filename.split("-")[1]
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}")

    pdf.ln()

    # Set header columns
    headers = list(df.columns)
    headers = [header.replace("_", " ").title() for header in headers]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(headers[0]), border=1, align="C")
    pdf.cell(w=60, h=8, txt=str(headers[1]), border=1, align="C")
    pdf.cell(w=40, h=8, txt=str(headers[2]), border=1, align="C")
    pdf.cell(w=30, h=8, txt=str(headers[3]), border=1, align="C")
    pdf.cell(w=30, h=8, txt=str(headers[4]), border=1, align="C")
    pdf.ln()

    # Set product rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1, align="C")
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1, align="C")
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1, align="R")
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1, align="R")
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, align="R")
        pdf.ln()

    pdf.output(f"pdfs/{filename}.pdf")
