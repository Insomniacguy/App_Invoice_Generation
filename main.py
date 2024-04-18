import pandas as pd
import glob
# from fpdf import FPDF   # importing FPDF class from fpdf module
import fpdf
import pathlib

# glob() returns a list of file paths that match the specified pattern.
filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)
print(type(filepaths))
# loading data(excel) into data frames using for loop for multiple excel files
for filepath in filepaths:
    # print(filepath)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # print(df)
    pdf = fpdf.FPDF(orientation='p', unit='mm', format='A4')

    filename = pathlib.Path(filepath).stem
    print(filename)
    print(type(filename))

    pdf.add_page()

    invoice_num, date = filename.split('-')
    # date = filename.split('-')[1]

    pdf.set_font("Times", size=12, style='B')
    pdf.cell(50, 12, txt=f"Invoice number: {invoice_num}", ln=1)

    pdf.set_font("Times", size=12, style='B')
    pdf.cell(50, 12, txt=f"Date: {date}")

    pdf.output(f"PDFS/{filename}.pdf")

# print(filepaths)
