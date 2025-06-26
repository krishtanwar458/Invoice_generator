import pandas as pd
from fpdf import FPDF
import glob

filepaths = glob.glob('/Users/krishtanwar/Desktop/Python/PDF generator/' \
'Invoice Generator/*.xlsx')

for i in filepaths:
    df = pd.read_excel(i, sheet_name='Sheet 1')
    # print(df)


