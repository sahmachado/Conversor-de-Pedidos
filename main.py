import pandas as pd
from conversor import conversor
import streamlit as st

st.title("Conversor")
st.write("Classificação de pedidos em atraso")
st.divider()
col1,col2 = st.columns(2)
with col1:
    df = st.file_uploader("Carregar arquivo CSV",type="csv",help="Carregue o arquivo CSV do relatório exportado do SAP")
if df is not None:
    with st.spinner("Convertendo arquivo..."):
        excel = conversor(df)
    if excel == None:
        st.error("Erro ao converter arquivo")
    else:
        st.success("Conversão concluída com sucesso!")
        st.write("Clique no botão abaixo para fazer o download do arquivo")
        btn = st.download_button(
            label= "Download",
            data = excel,
            file_name="Base Pedidos Atrasados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Faça download do arquivo excel convertido e classificado"
        )
        if btn:
            st.success("Arquivo baixado com sucesso!")