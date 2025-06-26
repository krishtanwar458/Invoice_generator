import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob('/Users/krishtanwar/Desktop/Python/PDF generator/' \
'Invoice Generator/*.xlsx')

for i in filepaths:
    df = pd.read_excel(i, sheet_name='Sheet 1')
    # print(df)
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', size=16, style='BI')
    filename = Path(i).stem
    invoice_num_date = filename.split('-')

    pdf.cell(w=50, h=8, txt=f'Inovice Number: {invoice_num_date[0]}', ln=1, align='L')
    pdf.cell(w=50, h=8, txt=f'Date: {invoice_num_date[1]}', ln=1, align='L')

    

    pdf.output(f'PDFs/{filename}.pdf')