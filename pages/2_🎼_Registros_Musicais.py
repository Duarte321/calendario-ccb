import mimetypes
import uuid
from datetime import date
import requests
import streamlit as st
from auth_utils import exigir_login, eh_admin, token, SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY

st.set_page_config(page_title="Registros Musicais", page_icon="🎼", layout="wide")
exigir_login()

def headers(prefer=None, content_type="application/json"):
    h={"apikey":SUPABASE_PUBLISHABLE_KEY,"Authorization":f"Bearer {token()}","Content-Type":content_type}
    if prefer: h["Prefer"]=prefer
    return h

def listar(tipo):
    r=requests.get(f"{SUPABASE_URL}/rest/v1/registros_musicais",headers=headers(),params={"tipo":f"eq.{tipo}","select":"*","order":"data.desc"},timeout=15)
    return r.json() if r.ok else []

def salvar_registro(payload):
    r=requests.post(f"{SUPABASE_URL}/rest/v1/registros_musicais",headers=headers("return=representation"),json=payload,timeout=15)
    r.raise_for_status()
    return r.json()[0]

def upload_arquivo(registro_id, arquivo, tipo):
    ext=arquivo.name.rsplit(".",1)[-1].lower()
    caminho=f"{registro_id}/{uuid.uuid4().hex}.{ext}"
    mime=arquivo.type or mimetypes.guess_type(arquivo.name)[0] or "application/octet-stream"
    h={"apikey":SUPABASE_PUBLISHABLE_KEY,"Authorization":f"Bearer {token()}","Content-Type":mime}
    r=requests.post(f"{SUPABASE_URL}/storage/v1/object/registros-musicais/{caminho}",headers=h,data=arquivo.getvalue(),timeout=30)
    r.raise_for_status()
    meta={"registro_id":registro_id,"tipo":tipo,"caminho":caminho,"nome_arquivo":arquivo.name}
    m=requests.post(f"{SUPABASE_URL}/rest/v1/registro_arquivos",headers=headers("return=representation"),json=meta,timeout=15)
    m.raise_for_status()

st.title("🎼 Registros Musicais")
st.caption("Arquivo histórico privado • Ensaios musicais e aulas do MSA")

tab1,tab2,tab3=st.tabs(["🎵 Ensaios","🎓 Aulas MSA","➕ Novo Registro"])

with tab1:
    st.subheader("Histórico de Ensaios")
    dados=listar("ensaio")
    if not dados: st.info("Nenhum ensaio cadastrado ainda.")
    for x in dados:
        with st.expander(f"🎵 {x['data']} • {x['titulo']} — {x['local']}"):
            st.write(x.get("resumo") or "Sem resumo.")
            c1,c2,c3=st.columns(3)
            c1.metric("Músicos",x.get("total_musicos") or 0)
            c2.metric("Organistas",x.get("total_organistas") or 0)
            c3.metric("Total",x.get("total_participantes") or 0)

with tab2:
    st.subheader("Histórico de Aulas do MSA")
    dados=listar("aula_msa")
    if not dados: st.info("Nenhuma aula cadastrada ainda.")
    for x in dados:
        with st.expander(f"🎓 {x['data']} • {x['titulo']} — {x['local']}"):
            if x.get("instrutor"): st.write(f"**Instrutor(a):** {x['instrutor']}")
            if x.get("assunto"): st.write(f"**Assunto:** {x['assunto']}")
            st.write(x.get("resumo") or "Sem resumo.")
            st.metric("Participantes",x.get("total_participantes") or 0)

with tab3:
    st.subheader("Novo Registro")
    if not eh_admin():
        st.warning("Somente administradores podem cadastrar registros.")
        st.stop()
    with st.form("novo_registro",clear_on_submit=True):
        tipo_ui=st.selectbox("Tipo",["Ensaio Musical","Aula do MSA"])
        data_reg=st.date_input("Data",value=date.today())
        local=st.text_input("Localidade")
        titulo=st.text_input("Título")
        instrutor=st.text_input("Responsável / Instrutor(a)")
        assunto=st.text_input("Assunto / Tema")
        resumo=st.text_area("Resumo / Observações")
        if tipo_ui=="Ensaio Musical":
            c1,c2=st.columns(2)
            musicos=c1.number_input("Total de músicos",min_value=0,step=1)
            organistas=c2.number_input("Total de organistas",min_value=0,step=1)
            participantes=int(musicos)+int(organistas)
        else:
            musicos=organistas=0
            participantes=st.number_input("Total de participantes",min_value=0,step=1)
        fotos=st.file_uploader("Fotos e documentos",type=["jpg","jpeg","png","webp","pdf"],accept_multiple_files=True)
        salvar=st.form_submit_button("💾 SALVAR REGISTRO",type="primary",use_container_width=True)
    if salvar:
        if not local.strip() or not titulo.strip():
            st.error("Informe a localidade e o título.")
        else:
            try:
                uid=(st.session_state.get("auth_user") or {}).get("id")
                payload={"tipo":"ensaio" if tipo_ui=="Ensaio Musical" else "aula_msa","data":data_reg.isoformat(),"local":local.strip(),"titulo":titulo.strip(),"instrutor":instrutor.strip() or None,"assunto":assunto.strip() or None,"resumo":resumo.strip() or None,"total_participantes":int(participantes),"total_musicos":int(musicos),"total_organistas":int(organistas),"criado_por":uid}
                registro=salvar_registro(payload)
                for arq in fotos or []:
                    upload_arquivo(registro["id"],arq,"foto")
                st.success("Registro salvo com sucesso! ✅")
                st.rerun()
            except Exception as e:
                st.error(f"Não foi possível salvar o registro: {e}")

st.divider()
st.caption("Agenda Musical • Região de Jaciara - MT")
