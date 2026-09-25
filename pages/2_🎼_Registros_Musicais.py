import streamlit as st
from datetime import date
import requests

st.set_page_config(page_title="Registros Musicais", page_icon="🎼", layout="wide")
st.title("🎼 Registros Musicais")
st.caption("Arquivo histórico de ensaios musicais e aulas do MSA")

SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "")

tab1, tab2, tab3 = st.tabs(["🎵 Ensaios", "🎓 Aulas MSA", "➕ Novo Registro"])

with tab1:
    st.subheader("Ensaios Musicais")
    st.info("Aqui ficará o histórico dos ensaios, resumos e galerias de fotos.")

with tab2:
    st.subheader("Aulas do MSA")
    st.info("Aqui ficará o histórico das aulas, fotos e listas de presença.")

with tab3:
    st.subheader("Novo Registro")
    tipo = st.selectbox("Tipo", ["Ensaio Musical", "Aula do MSA"])
    data = st.date_input("Data", value=date.today())
    local = st.text_input("Localidade")
    titulo = st.text_input("Título")
    if tipo == "Aula do MSA":
        instrutor = st.text_input("Instrutor(a)")
        assunto = st.text_input("Assunto da aula")
    resumo = st.text_area("Resumo / Observações")
    fotos = st.file_uploader("Fotos e documentos", type=["jpg","jpeg","png","webp","pdf"], accept_multiple_files=True)
    st.caption("O cadastro definitivo será liberado para administradores autenticados.")
    st.button("💾 Salvar registro", disabled=True)

st.divider()
st.caption("Agenda Musical • Região de Jaciara - MT")
