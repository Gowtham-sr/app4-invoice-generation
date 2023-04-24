import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
# glob creates a list of input data

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    # Path object identifies the given path & stem attribute returns the filename without file extension.
    invoice, date = filename.split("-")
    # split creates a list.

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice} ", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)
    pdf.ln()

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = df.columns
    # Returns index objects like List type. You can iterate it over.
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=68, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=27, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=68, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=27, h=8, txt=str(row["total_price"]), border=1, ln=1)
        # row["  "] returns an integer, but txt of fpdf library expects a string. So only converting

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    total = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=68, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=27, h=8, txt=str(total), border=1, ln=1)
    pdf.ln()

    # Add total sum sentence
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=33, h=8, txt=f"Way of Python")
    pdf.image("pythonhow.jpg", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
