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
    pdf.cell(w=50, h=8, txt=f"Date {invoice_date}", ln=1)

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
        pdf.cell(w=30, h=8, txt="$ {:.2f}".format(row["price_per_unit"]), border=1, align="R")
        pdf.cell(w=30, h=8, txt="$ {:.2f}".format(row["total_price"]), border=1, align="R")
        pdf.ln()

    # Order total Section
    total_amount_owed = df["total_price"].sum()
    total_formatted = "${:.2f}".format(total_amount_owed)

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="")
    pdf.cell(w=60, h=8, txt="")
    pdf.cell(w=40, h=8, txt="")
    pdf.cell(w=30, h=8, txt="TOTAL", align="R")
    pdf.set_fill_color(230, 230, 230)
    pdf.cell(w=30, h=8, txt=total_formatted, border=1, align="R", fill=True)
    pdf.ln()

    # Total Sum output
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_formatted}", ln=1)

    # Company Icon
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=20, h=8, txt="Owenitt LLC", ln=1)
    pdf.image(w=10, name="images/icon.jpg")
    pdf.ln(2)
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=0, txt='"Own it with Owenitt"')

    pdf.output(f"pdfs/{filename}.pdf")
