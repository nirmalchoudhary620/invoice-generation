import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Get all Excel files from the invoices folder
filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Extract invoice number and date from filename
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Create PDF document
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Add invoice number and date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Read Excel data
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns = [col.replace("_", " ").title() for col in df.columns]

    # Add table header
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add table rows
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add total row
    total_sum = df["total_price"].sum()
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=65, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total summary and footer
    pdf.set_text_color(0, 0, 0)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {str(total_sum)}", ln=1)

    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    # Save the PDF
    pdf.output(f"PDFs/{filename}.pdf")


