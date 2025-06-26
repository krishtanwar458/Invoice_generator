import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob('/Users/krishtanwar/Desktop/Python/PDF generator/' \
'Invoice Generator/*.xlsx')

for i in filepaths:
    # print(df)
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', size=16, style='BI')
    filename = Path(i).stem
    invoice_num_date = filename.split('-')

    pdf.cell(w=50, h=8, txt=f'Inovice Number: {invoice_num_date[0]}', ln=1, align='L')
    pdf.cell(w=50, h=8, txt=f'Date: {invoice_num_date[1]}', ln=1, align='L')

    df = pd.read_excel(i, sheet_name='Sheet 1')

    # Add header
    # type(df.columns)
    name = list(df.columns)
    pdf.set_font(family='Times', size=8, style='BI')
    pdf.cell(w=30, h=8, txt=name[0].title().replace('_', ' '), ln=0, border=1, align='C')
    pdf.cell(w=70, h=8, txt=name[1].title().replace('_', ' '), ln=0, border=1, align='C')
    pdf.cell(w=30, h=8, txt=name[2].title().replace('_', ' '), ln=0, border=1, align='C')
    pdf.cell(w=30, h=8, txt=name[3].title().replace('_', ' '), ln=0, border=1, align='C')
    pdf.cell(w=30, h=8, txt=name[4].title().replace('_', ' '), ln=1, border=1, align='C')

    # Add data
    total = []
    for i, row in df.iterrows():
        pdf.set_font(family='Times', size=10, style='')
        pdf.cell(w=30, h=8, txt=str(row['product_id']), ln=0, align='L', border=1)
        pdf.cell(w=70, h=8, txt=row['product_name'], ln=0, align='L', border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), ln=0, align='R', border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), ln=0, align='R', border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), ln=1, align='R', border=1)
        total.append(float(row['total_price']))

    pdf.cell(w=30, h=8, txt=(''), ln=0, align='L', border=1)
    pdf.cell(w=70, h=8, txt=(''), ln=0, align='L', border=1)
    pdf.cell(w=30, h=8, txt=(''), ln=0, align='R', border=1)
    pdf.cell(w=30, h=8, txt=(''), ln=0, align='R', border=1)

    full_amount = 0.0
    for i in total:
        full_amount = full_amount + i
    
    pdf.cell(w=30, h=8, txt=str(full_amount), ln=1, align='R', border=1)

    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=50, h=8, txt=f'The total amount due is {full_amount}', ln=1, align='R', border=0)
    pdf.cell(w=25, h=8, txt='PythonHow', ln=0, align='R', border=0)
    pdf.set_font(family='Times', size=14, style='B')
    pdf.image('pythonhow.png', w=10, h=10)
        
    pdf.output(f'PDFs/{filename}.pdf')