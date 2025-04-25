import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_number, invoice_date = filename.split('-')

    pdf.set_font(family='Times', style='B', size=24)
    pdf.cell(w=50, h=8, txt=f"Invoice number. {invoice_number}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    pdf.ln(10)
    columns = list(df.columns)
    columns = [item.replace('_', ' ').title() for item in columns]
    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)


    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=12)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=60, h=8, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=40, h=8, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=30, h=8, txt=f"{row['total_price']}", border=1, ln=1)



    pdf.output(f"PDFs/{filename}.pdf")