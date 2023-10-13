import pandas as pd

df = pd.read_excel(
    io="mes.xlsx",
    engine="openpyxl",
    sheet_name= "mes",
    skiprows=0,
    usecols="A:AA",
    nrows=10000
)
