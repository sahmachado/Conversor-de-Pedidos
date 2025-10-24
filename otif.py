import streamlit as st
from classificacao_otif import inicio # importa tua função que gera o Excel

st.write("Classificador otif")

st.write("Carregue os 3 arquivos CSV exportados do SAP: **ME80FN**, **ME2N** e **YBMM009**")

col1,col2 = st.columns(2)
with col1:
    arquivos = st.file_uploader(
    "Carregar arquivos CSV",
    type="csv",
    accept_multiple_files=True,
    help="Selecione os três relatórios CSV (ME80FN, ME2N e YB)"
)

if arquivos:
    # Dicionário pra identificar automaticamente
    me80fn, me2n, yb = None, None, None

    for arquivo in arquivos:
        nome = arquivo.name.lower()
        if "me80fn" in nome:
            me80fn = arquivo
        elif "me2n" in nome:
            me2n = arquivo
        elif "yb" in nome or "zmm_yb" in nome:
            yb = arquivo

    # Validação antes de gerar
    if not (me80fn and me2n and yb):
        st.warning("⚠️ Certifique-se de enviar **os três arquivos**: ME80FN, ME2N e YBMM009.")
    else:
        with st.spinner("Processando arquivos..."):
            buffer = inicio(me80fn, yb, me2n)
        st.success("✅ Base gerada com sucesso!")
        st.download_button(
            "Baixar arquivo Excel",
            data=buffer,
            file_name="Base_Otif.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
