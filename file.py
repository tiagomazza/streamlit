import plotly_express as px
import streamlit as st
import pandas as pd
import numpy as np
from dateutil import parser
import plotly.graph_objects as go

coeficienteDeDivisao = 11

df = pd.read_excel(
    io="mes.xlsx",
    engine="openpyxl",
    sheet_name= "mes",
    skiprows=0,
    usecols="A:J",
    nrows=4000
)

df= df.drop(columns=['A. BORGES DO AMARAL, Lda.'])

novos_nomes = {
    'Unnamed: 1': 'Data',
    'Unnamed: 2': 'CodigoCliente',
    'Unnamed: 3': 'Cliente',
    'Unnamed: 4': 'DescontoCliente',
    'Unnamed: 5': 'DescontoArtigo',
    'Unnamed: 6': 'NomeArtigo',
    'Unnamed: 7': 'ValorArtigo',
    'Unnamed: 8': 'Vendedor',
    'Unnamed: 9': 'CodigoVendedor'
}

df.rename(columns=novos_nomes, inplace=True)
df = df.dropna(subset=['ValorArtigo'])

df2 = pd.read_excel(
    io="2022.xlsx",
    engine="openpyxl",
    sheet_name= "2022",
    skiprows=0,
    usecols="A:G",
    nrows=40000
)

df2= df2.drop(columns=['A. BORGES DO AMARAL, Lda.'])

novos_nomes2 = {
    'Unnamed: 1': 'Data',
    'Unnamed: 2': 'CodigoCliente',
    'Unnamed: 3': 'Cliente',
    'Unnamed: 4': 'NomeArtigoLY',
    'Unnamed: 5': 'ValorArtigoLY',
    'Unnamed: 6': 'Vendedor',
}

df2.rename(columns=novos_nomes2, inplace=True)
df2 = df2.dropna(subset=['ValorArtigoLY'])

st.set_page_config(page_title="Sales",
                   page_icon=":bar_chart:",
                   layout="wide"
)

def limpar_e_converter(valor):
    if isinstance(valor, str):
        valor_limpo = valor.replace('\xa0', '')
        #.replace(',', '.')
        try:
            return float(valor_limpo)
        except ValueError:
            return None

    elif isinstance(valor, (float, int)):
        return valor  



def formatar_euro(valor):
    if isinstance(valor, (int, float)):
        return '{:,.2f}€'.format(valor)
    else:
        return str(valor)  

df = pd.concat([df, df2], join="outer", ignore_index=True)
df = df.iloc[1:]

df['ValorArtigo'] = df['ValorArtigo'].apply(limpar_e_converter)
df['ValorArtigoLY'] = df2['ValorArtigoLY'].apply(limpar_e_converter)

df['Data'] = pd.to_datetime(df['Data'], format='%d-%m-%Y', errors='coerce')
df['Mes_Ano'] = df['Data'].dt.strftime('%m-%Y')
df = df.sort_values(by='Cliente')

data = pd.read_excel('listagens.xlsx', sheet_name='Fornecedores')
data.loc[1:, 'Artigo'] = data['Artigo'][1:].astype(str)
dicionario_fornecedores = dict(zip(data['Artigo'], data['Fornecedor']))
df['Marca'] = df['NomeArtigo'].str[:3].map(dicionario_fornecedores)

def converter_para_numero(valor_str):
    partes = valor_str.split(" ")
    valor_sem_espaco = "".join(partes)
    try:
        return float(valor_sem_espaco)
    except ValueError:
        return None

df['ValorArtigo'] = df['ValorArtigo'].astype(str)
df['ValorArtigo'] = df['ValorArtigo'].apply(converter_para_numero)

#side bar

st.sidebar.header("Filtros de análise:")
vendedor = st.sidebar.multiselect(
    "selecione o vendedor:",
    options=df["Vendedor"].unique(),
    default=df["Vendedor"].unique()
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
df_selection = df[
    (df["Vendedor"].isin(vendedor)) &
    (df["Marca"].isin(marca)) &
    (df["Mes_Ano"].isin(mes_Ano)) &
    (df["Cliente"].isin(cliente))
]


#df_selection =df.query(
#   "Vendedor == @vendedor & Cliente==@cliente & Mes_Ano==@mes_Ano & Marca==@marca"
#)

# --- MAINPAGE ---

st.title(":bar_chart: Dashboard de vendas")
st.markdown("##")


df = df.iloc[1:]
df['ValorArtigo'] = df['ValorArtigo'].astype(str)
df['ValorArtigo'] = df['ValorArtigo'].apply(converter_para_numero)
total_sales = df_selection["ValorArtigo"].sum(skipna=True)

df['ValorArtigoLY'] = df['ValorArtigoLY'].astype(str)
df['ValorArtigoLY'] = df['ValorArtigoLY'].apply(converter_para_numero)


left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total de vendas:")
    st.subheader(f"{total_sales:,.2f}€")


with right_column:
    st.subheader("")

st.markdown("---")

# --- sales graphic ---

# Sales by client
sales_client = df_selection.groupby(by=["Cliente"])["ValorArtigo"].sum().reset_index()
sales_client = sales_client.sort_values(by="ValorArtigo", ascending=False)
sales_client["ValorArtigo"] = sales_client["ValorArtigo"].apply(formatar_euro)
sales_client = sales_client[::-1]
sales_client["Valor Líquido Formatado"] = sales_client["ValorArtigo"].apply(formatar_euro)

altura_desejada_por_cliente = 20  
altura_desejada = max(len(sales_client) * altura_desejada_por_cliente, 400)  

# Sales last year
sales_last = df.groupby(by=["Cliente"])["ValorArtigoLY"].sum().reset_index()
sales_last = sales_last.sort_values(by="ValorArtigoLY", ascending=True)
sales_last["ValorArtigoLY"] = sales_last["ValorArtigoLY"].apply(formatar_euro)
sales_last["Valor LíquidoLY"] = sales_last["ValorArtigoLY"].apply(lambda x: x / coeficienteDeDivisao if (isinstance(x, (int, float)) and coeficienteDeDivisao != 0) else x)
sales_last["Valor Líquido FormatadoLY"] = sales_last["Valor LíquidoLY"].apply(formatar_euro)

# -- grafico comparativo --

fig = go.Figure()

fig.add_trace(go.Bar(
    y=sales_last["Cliente"],
    x=sales_last["ValorArtigoLY"],
    text=sales_last["Valor Líquido FormatadoLY"], 
    name="Meta",
    orientation='h',
    marker=dict(color='red'),  
    width=0.5
))

fig.add_trace(go.Bar(
    y=sales_client["Cliente"],
    x=sales_client["ValorArtigo"],
      text=sales_client["Valor Líquido Formatado"], 
    name="Valor atual",
    orientation='h',
    marker=dict(color='blue'),  
    width=0.5
))

fig.update_layout(
    title="Comparativo de vendas",
    xaxis_title="Valores",
    yaxis_title="Cliente",
    barmode="overlay",
    width=1000,
    height=len(sales_last) * 20
)
fig.update_layout(plot_bgcolor="rgba(0,0,0,0)")

st.plotly_chart(fig)

hide_st_style = """
    <style>
    footer {visibility: hidden;}
    </style>
    """
st.markdown(hide_st_style, unsafe_allow_html=True)

