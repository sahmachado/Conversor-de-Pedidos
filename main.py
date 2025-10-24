import streamlit as st
from datetime import datetime




@st.cache_data
def cache():
    pass

pg = st.navigation(
    {'Conversores:':[
                st.Page('conversor.py',title='Pedidos em Atraso'),
                st.Page('otif.py',title='On Time In Full'),
                # st.Page('setor.py',title='Setor')
            ],
    }
)   
st.set_page_config(layout="wide")
st.title('Classificações de arquivos',help='Selecione a opção na barra de menu')
pg.run()