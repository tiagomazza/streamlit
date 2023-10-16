
import plotly_express as px
import streamlit as st
import pandas as pd
import numpy as np
from dateutil import parser
import plotly.graph_objects as go
import streamlit as st

df = pd.read_excel(
    io="ano.xlsx",
    engine="openpyxl",
    sheet_name= "Sheet1",
    skiprows=0,
    usecols="A:AC",
    nrows=40000
)
df = df.iloc[10:]
df = df.dropna(axis=1, how='all')
df = df.dropna(how='all')

novos_nomes = df.iloc[0]
df.columns = novos_nomes
df = df[1:]
df.reset_index(drop=True, inplace=True)

df2 = pd.read_excel(
    io="mes.xlsx",
    engine="openpyxl",
    sheet_name= "Sheet1",
    skiprows=0,
    usecols="A:AC",
    nrows=10000
)

df2 = df2.iloc[10:]
df2 = df2.dropna(axis=1, how='all')
df2 = df2.dropna(how='all')

novos_nomes = df2.iloc[0]
df2.columns = novos_nomes
df2 = df[1:]
df2.reset_index(drop=True, inplace=True)
#valor a ser dividido o anual do ano passado afim de uma média mensal.
fator_de_divisao = 11

st.set_page_config(page_title="Sales",
                   page_icon=":bar_chart:",
                   layout="wide"
)
#funçao de conversao dos clientes no formato 4 digitos
def format_string_to_4_digits(input_string):
    parts = input_string.split(".")
    formatted_string = parts[0]
    while len(formatted_string) < 4:
        formatted_string = "0" + formatted_string
    return formatted_string

def formatar_euro(valor):
    return '{:,.2f}€'.format(valor)

df ['Cliente'] = df.apply(lambda row:row['Data'] if row['Artigo'] != '' else None, axis=1)
df['Cliente'] = pd.to_numeric(df['Cliente'], errors='coerce')
df['Cliente'] = df['Cliente'].apply(lambda x: x if not pd.isna(x) else np.nan).ffill()
df['Cliente'] = df['Cliente'].astype(str)
df = df[~(df['Data'] == 'Total Cliente')]
df.dropna(subset=['Valor Líquido'], inplace=True)
df['Data'] = pd.to_datetime(df['Data'])
df['Mes_Ano'] = df['Data'].dt.strftime('%m-%Y')

df = df.sort_values(by='Cliente')

data = pd.read_excel('listagens.xlsx', sheet_name='Fornecedores')
data.loc[1:, 'Artigo'] = data['Artigo'][1:].astype(str)
dicionario_fornecedores = dict(zip(data['Artigo'], data['Fornecedor']))
df['Marca'] = df['Artigo'].str[:3].map(dicionario_fornecedores)

data = pd.read_excel('listagens.xlsx', sheet_name='Clientes')

data.loc[1:, 'Vendedor'] = data['Vendedor'][1:].astype(str)
data['Cliente'] = data['Cliente'].astype(str).apply(format_string_to_4_digits)
dicionario_clientes = dict(zip(data['Cliente'], data['Vendedor']))
df['Cliente'] = df['Cliente'].apply(format_string_to_4_digits)
df['Vendedor'] = df['Cliente'].str[:4].map(dicionario_clientes)



#---side bar
st.sidebar.header("Menu")
vendedores_disponiveis = df["Vendedor"].dropna().unique()
vendedor = st.sidebar.multiselect(
    "selecione o vendedor:",
    options=vendedores_disponiveis.tolist(),
    default=vendedores_disponiveis.tolist()
)

marca = st.sidebar.multiselect(
    "selecione a Marca",
    options=df["Marca"].unique(),
    default=df["Marca"].unique()
)
mes_Ano = st.sidebar.multiselect(
    "selecione o Mês Ano",
    options=df["Mes_Ano"].unique(),
    default=df["Mes_Ano"].unique()

)

cliente = st.sidebar.multiselect(
    "selecione o Cliente:",
    options=df["Cliente"].unique(),
    default=df["Cliente"].unique()
)
df_selection =df.query(
    "Vendedor == @vendedor & Cliente==@cliente & Mes_Ano==@mes_Ano & Marca==@marca"
)

# --- MAINPAGE ---


st.title(":bar_chart: Dashboard de vendas")
st.markdown("##")


total_sales = int(df_selection["Valor Líquido"].sum())
#average_reating = round(df_selection["Valor Líquido"].mean(),1)
#star_rating = ":star:" * int(round(average_reating, 0))
#average_sales_by_transaction = round(df_selection["Valor Líquido"].mean(),2)

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total de vendas:")
    st.subheader(f"{total_sales:,.2f}€")

#with middle_column:
 #   st.subheader("Avaliações:")
 #   st.subheader(f"{average_reating} {star_rating}")


with right_column:
    st.subheader("")
 #   st.subheader(f"€{average_sales_by_transaction:,}")

st.markdown("---")

# --- sales graphic ---
altura_desejada_por_cliente = 50  # altura em pixel
# sales by product line
df_selection["Valor Líquido"] = pd.to_numeric(df_selection["Valor Líquido"], errors="coerce")

sales_by_product_line = (
    df_selection.groupby(by=["Marca"])["Valor Líquido"].sum().reset_index()
)

sales_by_product_line = sales_by_product_line.sort_values(by="Valor Líquido", ascending=True)
sales_by_product_line["Valor Líquido Formatado"] = sales_by_product_line["Valor Líquido"].apply(formatar_euro)

fig_product_sales = px.bar(
    sales_by_product_line,
    x="Valor Líquido",
    y="Marca",
    text="Valor Líquido Formatado",
    title="Vendas por marca",
    color="Valor Líquido",
    color_continuous_scale=px.colors.sequential.Plasma,  # Escolha uma escala de cores
    width=800,
    height=1200
)

#fig_product_sales.update_traces(marker=dict(line=dict(width=2, color='DarkSlateGrey')))  # Adicione uma borda às barras

fig_product_sales.update_layout(plot_bgcolor="rgba(0,0,0,0)")
fig_product_sales.update_layout(yaxis_title="Marca", xaxis_title="Valor Líquido")
fig_product_sales.update_coloraxes(showscale=False)
st.plotly_chart(fig_product_sales)

# Sales by client
sales_client = df_selection.groupby(by=["Cliente"])["Valor Líquido"].sum().reset_index()
sales_client = sales_client.sort_values(by="Valor Líquido", ascending=True)
sales_client["Valor Líquido Formatado"] = sales_client["Valor Líquido"].apply(formatar_euro)

altura_desejada_por_cliente = 20  # Defina a altura desejada por cliente em pixels
altura_desejada = max(len(sales_client) * altura_desejada_por_cliente, 400)  # Defina uma altura mínima

fig_product_client = px.bar(
    sales_client,
    x="Valor Líquido",
    y="Cliente",
    text="Valor Líquido Formatado",
    title="Vendas por Cliente",
    color="Valor Líquido",
    color_continuous_scale=px.colors.sequential.Plasma,  # Escolha uma escala de cores
    width=800,
    height= altura_desejada
   
)

fig_product_client.update_layout(plot_bgcolor="rgba(0,0,0,0)")
fig_product_client.update_coloraxes(showscale=False)
st.plotly_chart(fig_product_client)

# -- grafico comparativo --
sales_client_per_month = sales_client 
sales_client_per_month["Valor Líquido por mês"] = sales_client_per_month["Valor Líquido"] / fator_de_divisao
fig = go.Figure()

sales_client_actual = df_selection.groupby(by=["Cliente"])["Valor Líquido"].sum().reset_index()
sales_client_actual = sales_client_actual.sort_values(by="Valor Líquido", ascending=True)
sales_client_actual["Valor Líquido Formatado"] = sales_client_actual["Valor Líquido"].apply(formatar_euro)

fig.add_trace(go.Bar(
    y=sales_client_per_month["Cliente"],
    x=sales_client_per_month["Valor Líquido por mês"],
    name="Meta",
    orientation='h',
    marker=dict(color='red'),  
    width=0.5

))

fig.add_trace(go.Bar(
    y=sales_client_actual["Cliente"],
    x=sales_client_actual["Valor Líquido Formatado"],
    name="Valor atual",
    orientation='h',
    marker=dict(color='blue'),  
    width=0.5
    
))

fig.update_layout(
    title="Gráfico de Barras Sobreposto Horizontal",
    xaxis_title="Valores",
    yaxis_title="Cliente",
    barmode="overlay",
    width=800,
    height=len(sales_client) * 15
)
fig.update_layout(plot_bgcolor="rgba(0,0,0,0)")

st.plotly_chart(fig)

hide_st_style = """
    <style>
    footer {visibility: hidden;}
    </style>
    """
st.markdown(hide_st_style, unsafe_allow_html=True)
