import streamlit as st
import pandas as pd

df = pd.read_excel(
    io="mes.xlsx",
    engine="openpyxl",
    sheet_name= "Sheet1",
    skiprows=0,
    usecols="A:AC",
    nrows=10000
)

df = df.iloc[10:]
df = df.dropna(axis=1, how='all')
df = df.dropna(how='all')

novos_nomes = df.iloc[0]
df.columns = novos_nomes
df = df[1:]
df.reset_index(drop=True, inplace=True)

print(df)
