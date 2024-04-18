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
    pdf.cell(50, 12, txt=f"Date: {date}", ln=1)

    #headers
    columns = df.columns  # index object is iterable no need to convert to list
    print(columns)
    columns = [item.replace('_', ' ').title() for item in columns]
    pdf.set_font(family='Times', size=12, style='B')
    pdf.set_text_color(0, 0, 0)
    pdf.cell(30, 12, txt=columns[0], border=1)
    pdf.cell(60, 12, txt=columns[1], border=1)
    pdf.cell(40, 12, txt=columns[2], border=1)
    pdf.cell(30, 12, txt=columns[3], border=1)
    pdf.cell(30, 12, txt=columns[4], border=1, ln=1)

    # rows
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=12, style='B')
        pdf.set_text_color(100, 100, 100)
        pdf.cell(30, 12, txt=str(row['product_id']), border=1)
        pdf.cell(60, 12, txt=str(row['product_name']), border=1)
        pdf.cell(40, 12, txt=str(row['amount_purchased']), border=1)
        pdf.cell(30, 12, txt=str(row['price_per_unit']), border=1)
        pdf.cell(30, 12, txt=str(row['total_price']), border=1, ln=1)

    total_price = df['total_price'].sum()
    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(30,12, border=1,)
    pdf.cell(60,12, border=1,)
    pdf.cell(40,12, border=1,)
    pdf.cell(30,12, border=1,)
    pdf.cell(0,12,txt=str(total_price), border=1, align='L', ln=1)

    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(15, 12, border=0, txt=f"The total price is {total_price}", ln=1)

    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(25, 12, border=0, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFS/{filename}.pdf")

# print(filepaths)
