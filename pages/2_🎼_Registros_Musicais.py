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

def listar_presencas(registro_id):
    r=requests.get(f"{SUPABASE_URL}/rest/v1/registro_presencas",headers=headers(),params={"registro_id":f"eq.{registro_id}","select":"*","order":"nome.asc"},timeout=15)
    return r.json() if r.ok else []

def salvar_presencas(registro_id,texto,categoria):
    for nome in [n.strip() for n in texto.splitlines() if n.strip()]:
        r=requests.post(f"{SUPABASE_URL}/rest/v1/registro_presencas",headers=headers(),json={"registro_id":registro_id,"nome":nome,"categoria":categoria},timeout=15)
        r.raise_for_status()

def listar_arquivos(registro_id):
    r=requests.get(f"{SUPABASE_URL}/rest/v1/registro_arquivos",headers=headers(),params={"registro_id":f"eq.{registro_id}","select":"*","order":"criado_em.asc"},timeout=15)
    return r.json() if r.ok else []

def baixar_arquivo(caminho):
    r=requests.get(f"{SUPABASE_URL}/storage/v1/object/authenticated/registros-musicais/{caminho}",headers={"apikey":SUPABASE_PUBLISHABLE_KEY,"Authorization":f"Bearer {token()}"},timeout=30)
    return r.content if r.ok else None

def atualizar_registro(registro_id,payload):
    r=requests.patch(f"{SUPABASE_URL}/rest/v1/registros_musicais",headers=headers("return=representation"),params={"id":f"eq.{registro_id}"},json=payload,timeout=15)
    r.raise_for_status()

def excluir_registro(registro_id):
    r=requests.delete(f"{SUPABASE_URL}/rest/v1/registros_musicais",headers=headers("return=representation"),params={"id":f"eq.{registro_id}"},timeout=15)
    r.raise_for_status()

def detalhes_registro(x):
    pres=listar_presencas(x["id"])
    if pres:
        st.markdown("**👥 Lista de presença**")
        for p in pres: st.write(f"• {p['nome']} — {p.get('categoria') or 'Participante'}")
    arquivos=listar_arquivos(x["id"])
    imagens=[]
    for arq in arquivos:
        dados=baixar_arquivo(arq["caminho"])
        if not dados: continue
        nome=arq.get("nome_arquivo") or "arquivo"
        if nome.lower().endswith((".jpg",".jpeg",".png",".webp")): imagens.append((dados,nome))
        else: st.download_button(f"📄 {nome}",dados,file_name=nome,key=f"dl_{arq['id']}")
    if imagens:
        st.markdown("**📸 Galeria de fotos**")
        cols=st.columns(3)
        for i,(img,nome) in enumerate(imagens): cols[i%3].image(img,caption=nome,use_container_width=True)
    if eh_admin():
        with st.expander("⚙️ Editar / Excluir"):
            novo_titulo=st.text_input("Título",x["titulo"],key=f"tit_{x['id']}")
            novo_local=st.text_input("Localidade",x["local"],key=f"loc_{x['id']}")
            novo_resumo=st.text_area("Resumo",x.get("resumo") or "",key=f"res_{x['id']}")
            c1,c2=st.columns(2)
            if c1.button("💾 Salvar alterações",key=f"save_{x['id']}"):
                atualizar_registro(x["id"],{"titulo":novo_titulo,"local":novo_local,"resumo":novo_resumo or None})
                st.success("Alterações salvas."); st.rerun()
            confirmar=c2.checkbox("Confirmar exclusão",key=f"conf_{x['id']}")
            if c2.button("🗑️ Excluir",key=f"del_{x['id']}",disabled=not confirmar):
                excluir_registro(x["id"]); st.success("Registro excluído."); st.rerun()

st.title("🎼 Registros Musicais")
st.caption("Arquivo histórico privado • Ensaios musicais e aulas do MSA")

secao_padrao=st.session_state.get("registro_secao","Ensaios")
opcoes=["Ensaios","Aulas MSA","Novo Registro"]
indice=opcoes.index(secao_padrao) if secao_padrao in opcoes else 0
secao=st.radio("Navegação",opcoes,index=indice,horizontal=True,label_visibility="collapsed")
st.session_state["registro_secao"]=secao

f1,f2,f3=st.columns(3)
ano_filtro=f1.selectbox("Ano",["Todos"]+list(range(date.today().year,2020,-1)))
mes_filtro=f2.selectbox("Mês",["Todos"]+list(range(1,13)))
local_filtro=f3.text_input("🔎 Localidade")

def filtrar(dados):
    saida=[]
    for x in dados:
        y,m=map(int,x["data"].split("-")[:2])
        if ano_filtro!="Todos" and y!=int(ano_filtro): continue
        if mes_filtro!="Todos" and m!=int(mes_filtro): continue
        if local_filtro and local_filtro.lower() not in x["local"].lower(): continue
        saida.append(x)
    return saida

if secao=="Ensaios":
    st.subheader("Histórico de Ensaios")
    dados=filtrar(listar("ensaio"))
    if not dados: st.info("Nenhum ensaio cadastrado ainda.")
    for x in dados:
        with st.expander(f"🎵 {x['data']} • {x['titulo']} — {x['local']}"):
            st.write(x.get("resumo") or "Sem resumo.")
            c1,c2,c3=st.columns(3)
            c1.metric("Músicos",x.get("total_musicos") or 0)
            c2.metric("Organistas",x.get("total_organistas") or 0)
            c3.metric("Total",x.get("total_participantes") or 0)
            detalhes_registro(x)

elif secao=="Aulas MSA":
    st.subheader("Histórico de Aulas do MSA")
    dados=filtrar(listar("aula_msa"))
    if not dados: st.info("Nenhuma aula cadastrada ainda.")
    for x in dados:
        with st.expander(f"🎓 {x['data']} • {x['titulo']} — {x['local']}"):
            if x.get("instrutor"): st.write(f"**Instrutor(a):** {x['instrutor']}")
            if x.get("assunto"): st.write(f"**Assunto:** {x['assunto']}")
            st.write(x.get("resumo") or "Sem resumo.")
            st.metric("Participantes",x.get("total_participantes") or 0)
            detalhes_registro(x)

elif secao=="Novo Registro":
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
        st.markdown("#### 👥 Presença")
        presencas=st.text_area("Nomes dos participantes (um por linha)",placeholder="Nome 1\nNome 2\nNome 3")
        categoria_presenca=st.selectbox("Categoria da lista",["Participante","Músico","Organista","Instrutor(a)"])
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
                if presencas.strip():
                    salvar_presencas(registro["id"],presencas,categoria_presenca)
                for arq in fotos or []:
                    upload_arquivo(registro["id"],arq,"foto")
                st.success("Registro salvo com sucesso! ✅")
                st.rerun()
            except Exception as e:
                st.error(f"Não foi possível salvar o registro: {e}")

st.divider()
st.caption("Agenda Musical • Região de Jaciara - MT")
