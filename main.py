import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()


