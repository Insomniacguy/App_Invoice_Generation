import pandas as pd
import glob

# glob() returns a list of file paths that match the specified pattern.
filepaths = glob.glob("Invoices/*.xlsx")

# loading data(excel) into data frames using for loop for multiple excel files
for filepath in filepaths:
    # print(filepath)
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)

# print(filepaths)
