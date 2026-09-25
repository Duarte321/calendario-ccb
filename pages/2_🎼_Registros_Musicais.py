import html
import mimetypes
import uuid
from datetime import date, datetime

import requests
import streamlit as st

from auth_utils import (
    SUPABASE_PUBLISHABLE_KEY,
    SUPABASE_URL,
    eh_admin,
    exigir_login,
    logout,
    token,
)

st.set_page_config(page_title="Registros Musicais", page_icon="🎼", layout="wide")
exigir_login()

st.markdown(
    """
    <style>
    [data-testid="stSidebar"]{display:none}
    .block-container{max-width:1320px;padding-top:3.8rem;padding-bottom:2rem}

    .rm-topbar{
        background:#ffffff;border:1px solid #e6ebf0;border-radius:16px;
        padding:10px 14px;margin:0 0 18px;
        box-shadow:0 8px 22px rgba(15,23,42,.05);
        color:#64748b;font-size:13px
    }
    .rm-user{
        text-align:center;color:#64748b;padding-top:10px;font-size:13px;
        white-space:nowrap;overflow:hidden;text-overflow:ellipsis
    }
    .rm-nav-note{font-size:11px;color:#94a3b8;text-align:center;margin-top:2px}
    div[data-testid="stHorizontalBlock"]:has(.rm-user){align-items:center}

    .rm-hero{
        background:linear-gradient(135deg,#06233b 0%,#0b3556 55%,#0e4268 100%);
        border:1px solid rgba(226,179,63,.28);border-radius:24px;
        padding:30px 32px;margin:8px 0 22px;color:white;
        box-shadow:0 18px 45px rgba(7,36,59,.16)
    }
    .rm-kicker{color:#f2c45e;font-size:12px;font-weight:800;letter-spacing:.12em;text-transform:uppercase}
    .rm-title{font-family:Georgia,serif;font-size:38px;font-weight:800;margin:7px 0 5px;color:#fff}
    .rm-sub{color:#dbe8f2;font-size:15px;max-width:760px;line-height:1.6}
    .rm-badge{
        display:inline-block;margin-top:14px;padding:6px 11px;border-radius:999px;
        background:rgba(255,255,255,.09);border:1px solid rgba(255,255,255,.13);
        color:#f8fafc;font-size:12px;font-weight:700
    }

    .rm-stat{
        background:#fff;border:1px solid #e7ecf2;border-radius:18px;
        padding:17px 18px;box-shadow:0 8px 24px rgba(15,23,42,.055);
        min-height:112px
    }
    .rm-stat-label{font-size:11px;font-weight:800;letter-spacing:.07em;color:#7c8798;text-transform:uppercase}
    .rm-stat-value{font-size:29px;font-weight:850;color:#082d49;margin-top:8px;line-height:1}
    .rm-stat-note{font-size:12px;color:#94a3b8;margin-top:7px}

    .rm-section-title{font-size:22px;font-weight:850;color:#082d49;margin:26px 0 12px}
    .rm-filter-head{font-size:13px;font-weight:800;color:#082d49;margin:0 0 8px}

    .rm-card{
        background:#fff;border:1px solid #e4eaf0;border-radius:20px;
        padding:20px 22px;margin:0 0 14px;
        box-shadow:0 10px 28px rgba(15,23,42,.055)
    }
    .rm-card:hover{border-color:#d6b355;box-shadow:0 14px 34px rgba(15,23,42,.08)}
    .rm-card-grid{display:grid;grid-template-columns:92px 1fr;gap:18px;align-items:start}
    .rm-date{
        background:linear-gradient(180deg,#f8fbfd,#f1f6f9);
        border:1px solid #e1e8ee;border-radius:16px;padding:13px 8px;text-align:center
    }
    .rm-date-day{font-size:28px;font-weight:900;color:#082d49;line-height:1}
    .rm-date-month{font-size:11px;font-weight:800;color:#8b6a13;margin-top:7px;letter-spacing:.07em}
    .rm-type{
        display:inline-block;background:#fff7df;color:#8b6500;border:1px solid #f3df9f;
        border-radius:999px;padding:4px 9px;font-size:10px;font-weight:850;letter-spacing:.06em
    }
    .rm-card-title{font-size:21px;font-weight:850;color:#082d49;margin:7px 0 4px}
    .rm-card-local{font-size:13px;color:#64748b}
    .rm-card-summary{font-size:13px;color:#475569;line-height:1.55;margin-top:10px}
    .rm-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:9px;margin-top:14px}
    .rm-mini{
        background:#f8fafc;border:1px solid #e9eef3;border-radius:13px;padding:10px 12px
    }
    .rm-mini-label{font-size:10px;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.06em}
    .rm-mini-value{font-size:19px;font-weight:900;color:#0b3556;margin-top:3px}

    .rm-empty{
        background:linear-gradient(180deg,#fff,#f7fafc);border:1px dashed #cdd8e2;
        border-radius:22px;text-align:center;padding:42px 24px;
        box-shadow:0 8px 24px rgba(15,23,42,.035)
    }
    .rm-empty-icon{font-size:40px}.rm-empty-title{font-size:21px;font-weight:850;color:#082d49;margin-top:8px}
    .rm-empty-text{color:#64748b;font-size:13px;margin-top:5px}

    .rm-form-head{
        background:#f8fbfd;border:1px solid #e3ebf1;border-radius:18px;
        padding:16px 18px;margin:8px 0 14px
    }
    .rm-form-title{font-size:18px;font-weight:850;color:#082d49}
    .rm-form-sub{font-size:12px;color:#64748b;margin-top:3px}

    div[data-testid="stRadio"] > div{gap:12px}
    div[data-testid="stMetric"]{
        background:#fff;border:1px solid #e7ecf2;border-radius:14px;padding:10px 14px
    }
    .stButton>button{border-radius:12px;font-weight:750}
    div[data-testid="stExpander"]{border:1px solid #e4eaf0;border-radius:14px;overflow:hidden}

    @media(max-width:800px){
        .rm-title{font-size:30px}
        .rm-hero{padding:24px 20px}
        .rm-card-grid{grid-template-columns:1fr}
        .rm-date{width:90px}
        .rm-metrics{grid-template-columns:1fr}
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Navegação ----------
st.markdown('<div class="rm-topbar">', unsafe_allow_html=True)
top1, top2, top3 = st.columns([1.45, 6.1, 1.45], gap="medium")
with top1:
    if st.button("← VOLTAR AO PAINEL", use_container_width=True, key="rm_back"):
        st.session_state["nav"] = "Painel"
        st.switch_page("app.py")
with top2:
    perfil = st.session_state.get("perfil") or {}
    nome = perfil.get("nome") or (st.session_state.get("auth_user") or {}).get("email", "Usuário")
    papel = "Administrador" if eh_admin() else "Usuário"
    st.markdown(
        f'<div class="rm-user">👤 {html.escape(str(nome))} &nbsp;•&nbsp; {papel}</div>'
        '<div class="rm-nav-note">Registros Musicais</div>',
        unsafe_allow_html=True,
    )
with top3:
    if st.button("🚪 SAIR", use_container_width=True, key="rm_logout"):
        logout()
        st.session_state["nav"] = "Painel"
        st.switch_page("app.py")
st.markdown('</div>', unsafe_allow_html=True)


# ---------- Supabase ----------
def headers(prefer=None, content_type="application/json"):
    h = {
        "apikey": SUPABASE_PUBLISHABLE_KEY,
        "Authorization": f"Bearer {token()}",
        "Content-Type": content_type,
    }
    if prefer:
        h["Prefer"] = prefer
    return h


def listar(tipo):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/registros_musicais",
        headers=headers(),
        params={"tipo": f"eq.{tipo}", "select": "*", "order": "data.desc"},
        timeout=15,
    )
    return r.json() if r.ok else []


def salvar_registro(payload):
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/registros_musicais",
        headers=headers("return=representation"),
        json=payload,
        timeout=15,
    )
    r.raise_for_status()
    return r.json()[0]


def upload_arquivo(registro_id, arquivo, tipo):
    ext = arquivo.name.rsplit(".", 1)[-1].lower()
    caminho = f"{registro_id}/{uuid.uuid4().hex}.{ext}"
    mime = arquivo.type or mimetypes.guess_type(arquivo.name)[0] or "application/octet-stream"
    h = {
        "apikey": SUPABASE_PUBLISHABLE_KEY,
        "Authorization": f"Bearer {token()}",
        "Content-Type": mime,
    }
    r = requests.post(
        f"{SUPABASE_URL}/storage/v1/object/registros-musicais/{caminho}",
        headers=h,
        data=arquivo.getvalue(),
        timeout=30,
    )
    r.raise_for_status()
    meta = {
        "registro_id": registro_id,
        "tipo": tipo,
        "caminho": caminho,
        "nome_arquivo": arquivo.name,
    }
    m = requests.post(
        f"{SUPABASE_URL}/rest/v1/registro_arquivos",
        headers=headers("return=representation"),
        json=meta,
        timeout=15,
    )
    m.raise_for_status()


def listar_presencas(registro_id):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/registro_presencas",
        headers=headers(),
        params={"registro_id": f"eq.{registro_id}", "select": "*", "order": "nome.asc"},
        timeout=15,
    )
    return r.json() if r.ok else []


def salvar_presencas(registro_id, texto, categoria):
    for pessoa in [n.strip() for n in texto.splitlines() if n.strip()]:
        r = requests.post(
            f"{SUPABASE_URL}/rest/v1/registro_presencas",
            headers=headers(),
            json={"registro_id": registro_id, "nome": pessoa, "categoria": categoria},
            timeout=15,
        )
        r.raise_for_status()


def listar_arquivos(registro_id):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/registro_arquivos",
        headers=headers(),
        params={"registro_id": f"eq.{registro_id}", "select": "*", "order": "criado_em.asc"},
        timeout=15,
    )
    return r.json() if r.ok else []


def baixar_arquivo(caminho):
    r = requests.get(
        f"{SUPABASE_URL}/storage/v1/object/authenticated/registros-musicais/{caminho}",
        headers={"apikey": SUPABASE_PUBLISHABLE_KEY, "Authorization": f"Bearer {token()}"},
        timeout=30,
    )
    return r.content if r.ok else None


def atualizar_registro(registro_id, payload):
    r = requests.patch(
        f"{SUPABASE_URL}/rest/v1/registros_musicais",
        headers=headers("return=representation"),
        params={"id": f"eq.{registro_id}"},
        json=payload,
        timeout=15,
    )
    r.raise_for_status()


def excluir_registro(registro_id):
    r = requests.delete(
        f"{SUPABASE_URL}/rest/v1/registros_musicais",
        headers=headers("return=representation"),
        params={"id": f"eq.{registro_id}"},
        timeout=15,
    )
    r.raise_for_status()


# ---------- Utilidades visuais ----------
MESES_CURTOS = {
    1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR", 5: "MAI", 6: "JUN",
    7: "JUL", 8: "AGO", 9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ",
}


def data_partes(data_iso):
    try:
        d = datetime.strptime(data_iso, "%Y-%m-%d")
        return d.day, MESES_CURTOS[d.month], d.year
    except Exception:
        return "--", "", ""


def detalhes_registro(x):
    pres = listar_presencas(x["id"])
    if pres:
        st.markdown("#### 👥 Lista de presença")
        cols = st.columns(2)
        for i, p in enumerate(pres):
            categoria = p.get("categoria") or "Participante"
            cols[i % 2].markdown(f"**{html.escape(p['nome'])}**  \\n{html.escape(categoria)}")

    arquivos = listar_arquivos(x["id"])
    imagens = []
    documentos = []
    for arq in arquivos:
        dados = baixar_arquivo(arq["caminho"])
        if not dados:
            continue
        arq_nome = arq.get("nome_arquivo") or "arquivo"
        if arq_nome.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
            imagens.append((dados, arq_nome))
        else:
            documentos.append((dados, arq_nome, arq["id"]))

    if imagens:
        st.markdown("#### 📸 Galeria de fotos")
        cols = st.columns(3)
        for i, (img, legenda) in enumerate(imagens):
            cols[i % 3].image(img, caption=legenda, use_container_width=True)

    if documentos:
        st.markdown("#### 📄 Documentos")
        for dados, arq_nome, arq_id in documentos:
            st.download_button(
                f"Baixar {arq_nome}",
                dados,
                file_name=arq_nome,
                key=f"dl_{arq_id}",
            )

    if eh_admin():
        with st.expander("⚙️ Editar ou excluir este registro"):
            novo_titulo = st.text_input("Título", x["titulo"], key=f"tit_{x['id']}")
            novo_local = st.text_input("Localidade", x["local"], key=f"loc_{x['id']}")
            novo_resumo = st.text_area("Resumo", x.get("resumo") or "", key=f"res_{x['id']}")

            c1, c2 = st.columns(2)
            if c1.button("💾 Salvar alterações", key=f"save_{x['id']}", use_container_width=True):
                atualizar_registro(
                    x["id"],
                    {"titulo": novo_titulo, "local": novo_local, "resumo": novo_resumo or None},
                )
                st.success("Alterações salvas.")
                st.rerun()

            confirmar = c2.checkbox("Confirmar exclusão", key=f"conf_{x['id']}")
            if c2.button(
                "🗑️ Excluir registro",
                key=f"del_{x['id']}",
                disabled=not confirmar,
                use_container_width=True,
            ):
                excluir_registro(x["id"])
                st.success("Registro excluído.")
                st.rerun()


def card_ensaio(x):
    dia, mes, ano = data_partes(x["data"])
    titulo = html.escape(str(x.get("titulo") or "Ensaio Musical"))
    local = html.escape(str(x.get("local") or "Local não informado"))
    resumo = html.escape(str(x.get("resumo") or "Sem resumo informado."))
    musicos = int(x.get("total_musicos") or 0)
    organistas = int(x.get("total_organistas") or 0)
    total = int(x.get("total_participantes") or 0)

    st.markdown(
        f"""
        <div class="rm-card">
          <div class="rm-card-grid">
            <div class="rm-date">
              <div class="rm-date-day">{dia}</div>
              <div class="rm-date-month">{mes} {ano}</div>
            </div>
            <div>
              <span class="rm-type">ENSAIO MUSICAL</span>
              <div class="rm-card-title">{titulo}</div>
              <div class="rm-card-local">📍 {local}</div>
              <div class="rm-card-summary">{resumo}</div>
              <div class="rm-metrics">
                <div class="rm-mini"><div class="rm-mini-label">Músicos</div><div class="rm-mini-value">{musicos}</div></div>
                <div class="rm-mini"><div class="rm-mini-label">Organistas</div><div class="rm-mini-value">{organistas}</div></div>
                <div class="rm-mini"><div class="rm-mini-label">Total</div><div class="rm-mini-value">{total}</div></div>
              </div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.expander("Ver fotos, presença e detalhes"):
        detalhes_registro(x)


def card_aula(x):
    dia, mes, ano = data_partes(x["data"])
    titulo = html.escape(str(x.get("titulo") or "Aula do MSA"))
    local = html.escape(str(x.get("local") or "Local não informado"))
    resumo = html.escape(str(x.get("resumo") or "Sem resumo informado."))
    instrutor = html.escape(str(x.get("instrutor") or "Não informado"))
    assunto = html.escape(str(x.get("assunto") or "Não informado"))
    total = int(x.get("total_participantes") or 0)

    st.markdown(
        f"""
        <div class="rm-card">
          <div class="rm-card-grid">
            <div class="rm-date">
              <div class="rm-date-day">{dia}</div>
              <div class="rm-date-month">{mes} {ano}</div>
            </div>
            <div>
              <span class="rm-type">AULA DO MSA</span>
              <div class="rm-card-title">{titulo}</div>
              <div class="rm-card-local">📍 {local}</div>
              <div class="rm-card-summary"><b>Instrutor(a):</b> {instrutor} &nbsp; • &nbsp; <b>Assunto:</b> {assunto}</div>
              <div class="rm-card-summary">{resumo}</div>
              <div class="rm-metrics" style="grid-template-columns:1fr">
                <div class="rm-mini"><div class="rm-mini-label">Participantes</div><div class="rm-mini-value">{total}</div></div>
              </div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.expander("Ver fotos, presença e detalhes"):
        detalhes_registro(x)


# ---------- Cabeçalho ----------
st.markdown(
    """
    <div class="rm-hero">
      <div class="rm-kicker">Arquivo histórico privado</div>
      <div class="rm-title">🎼 Registros Musicais</div>
      <div class="rm-sub">Consulte ensaios, aulas do MSA, fotos, presença e documentos em um único ambiente organizado.</div>
      <span class="rm-badge">Região de Jaciara - MT</span>
    </div>
    """,
    unsafe_allow_html=True,
)

secao_padrao = st.session_state.get("registro_secao", "Ensaios")
opcoes = ["Ensaios", "Aulas MSA", "Novo Registro"]
indice = opcoes.index(secao_padrao) if secao_padrao in opcoes else 0
secao = st.radio("Navegação", opcoes, index=indice, horizontal=True, label_visibility="collapsed")
st.session_state["registro_secao"] = secao

# Dados para indicadores e filtros
todos_ensaios = listar("ensaio")
todas_aulas = listar("aula_msa")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="rm-stat"><div class="rm-stat-label">Ensaios registrados</div><div class="rm-stat-value">{len(todos_ensaios)}</div><div class="rm-stat-note">Histórico musical</div></div>',
        unsafe_allow_html=True,
    )
with c2:
    soma_musicos = sum(int(x.get("total_musicos") or 0) for x in todos_ensaios)
    st.markdown(
        f'<div class="rm-stat"><div class="rm-stat-label">Músicos registrados</div><div class="rm-stat-value">{soma_musicos}</div><div class="rm-stat-note">Somatório dos ensaios</div></div>',
        unsafe_allow_html=True,
    )
with c3:
    soma_organistas = sum(int(x.get("total_organistas") or 0) for x in todos_ensaios)
    st.markdown(
        f'<div class="rm-stat"><div class="rm-stat-label">Organistas registrados</div><div class="rm-stat-value">{soma_organistas}</div><div class="rm-stat-note">Somatório dos ensaios</div></div>',
        unsafe_allow_html=True,
    )
with c4:
    st.markdown(
        f'<div class="rm-stat"><div class="rm-stat-label">Aulas do MSA</div><div class="rm-stat-value">{len(todas_aulas)}</div><div class="rm-stat-note">Aulas arquivadas</div></div>',
        unsafe_allow_html=True,
    )

st.markdown('<div class="rm-section-title">🔎 Localizar registros</div>', unsafe_allow_html=True)
f1, f2, f3 = st.columns(3)
ano_filtro = f1.selectbox("Ano", ["Todos"] + list(range(date.today().year, 2020, -1)))
mes_filtro = f2.selectbox(
    "Mês",
    ["Todos"] + list(range(1, 13)),
    format_func=lambda x: x if x == "Todos" else f"{x:02d} • {MESES_CURTOS[x]}",
)
local_filtro = f3.text_input("Localidade", placeholder="Ex.: Jaciara - Central")


def filtrar(dados):
    saida = []
    for x in dados:
        y, m = map(int, x["data"].split("-")[:2])
        if ano_filtro != "Todos" and y != int(ano_filtro):
            continue
        if mes_filtro != "Todos" and m != int(mes_filtro):
            continue
        if local_filtro and local_filtro.lower() not in (x.get("local") or "").lower():
            continue
        saida.append(x)
    return saida


# ---------- Conteúdo ----------
if secao == "Ensaios":
    dados = filtrar(todos_ensaios)
    st.markdown(
        f'<div class="rm-section-title">🎵 Histórico de Ensaios <span style="font-size:13px;color:#94a3b8;font-weight:600">• {len(dados)} registro(s)</span></div>',
        unsafe_allow_html=True,
    )

    if not dados:
        st.markdown(
            """
            <div class="rm-empty">
              <div class="rm-empty-icon">🎼</div>
              <div class="rm-empty-title">Nenhum ensaio encontrado</div>
              <div class="rm-empty-text">Os ensaios cadastrados aparecerão aqui com fotos, presença e informações organizadas.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if eh_admin():
            st.markdown("")
            if st.button("➕ CADASTRAR PRIMEIRO ENSAIO", type="primary", use_container_width=True):
                st.session_state["registro_secao"] = "Novo Registro"
                st.rerun()
    else:
        for x in dados:
            card_ensaio(x)

elif secao == "Aulas MSA":
    dados = filtrar(todas_aulas)
    st.markdown(
        f'<div class="rm-section-title">🎓 Histórico de Aulas do MSA <span style="font-size:13px;color:#94a3b8;font-weight:600">• {len(dados)} registro(s)</span></div>',
        unsafe_allow_html=True,
    )

    if not dados:
        st.markdown(
            """
            <div class="rm-empty">
              <div class="rm-empty-icon">🎓</div>
              <div class="rm-empty-title">Nenhuma aula encontrada</div>
              <div class="rm-empty-text">As aulas do MSA aparecerão aqui com tema, instrutor, participantes e arquivos.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        for x in dados:
            card_aula(x)

elif secao == "Novo Registro":
    st.markdown(
        """
        <div class="rm-form-head">
          <div class="rm-form-title">➕ Novo Registro Musical</div>
          <div class="rm-form-sub">Cadastre um ensaio ou uma aula do MSA e anexe fotos, documentos e participantes.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not eh_admin():
        st.warning("Somente administradores podem cadastrar registros.")
        st.stop()

    with st.form("novo_registro", clear_on_submit=True):
        tipo_ui = st.selectbox("Tipo de registro", ["Ensaio Musical", "Aula do MSA"])

        r1, r2 = st.columns(2)
        data_reg = r1.date_input("Data", value=date.today())
        local = r2.text_input("Localidade", placeholder="Ex.: Jaciara - Central")

        titulo = st.text_input("Título", placeholder="Ex.: Ensaio Local")
        r3, r4 = st.columns(2)
        instrutor = r3.text_input("Responsável / Instrutor(a)")
        assunto = r4.text_input("Assunto / Tema")
        resumo = st.text_area("Resumo / Observações", height=120)

        if tipo_ui == "Ensaio Musical":
            c1, c2 = st.columns(2)
            musicos = c1.number_input("Total de músicos", min_value=0, step=1)
            organistas = c2.number_input("Total de organistas", min_value=0, step=1)
            participantes = int(musicos) + int(organistas)
        else:
            musicos = 0
            organistas = 0
            participantes = st.number_input("Total de participantes", min_value=0, step=1)

        st.markdown("#### 👥 Presença")
        presencas = st.text_area(
            "Nomes dos participantes (um por linha)",
            placeholder="Nome 1\nNome 2\nNome 3",
        )
        categoria_presenca = st.selectbox(
            "Categoria da lista",
            ["Participante", "Músico", "Organista", "Instrutor(a)"],
        )

        st.markdown("#### 📎 Fotos e documentos")
        fotos = st.file_uploader(
            "Selecione os arquivos",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
        )

        salvar = st.form_submit_button(
            "💾 SALVAR REGISTRO",
            type="primary",
            use_container_width=True,
        )

    if salvar:
        if not local.strip() or not titulo.strip():
            st.error("Informe a localidade e o título.")
        else:
            try:
                uid = (st.session_state.get("auth_user") or {}).get("id")
                payload = {
                    "tipo": "ensaio" if tipo_ui == "Ensaio Musical" else "aula_msa",
                    "data": data_reg.isoformat(),
                    "local": local.strip(),
                    "titulo": titulo.strip(),
                    "instrutor": instrutor.strip() or None,
                    "assunto": assunto.strip() or None,
                    "resumo": resumo.strip() or None,
                    "total_participantes": int(participantes),
                    "total_musicos": int(musicos),
                    "total_organistas": int(organistas),
                    "criado_por": uid,
                }
                registro = salvar_registro(payload)

                if presencas.strip():
                    salvar_presencas(registro["id"], presencas, categoria_presenca)

                for arq in fotos or []:
                    upload_arquivo(registro["id"], arq, "foto")

                st.success("Registro salvo com sucesso! ✅")
                st.session_state["registro_secao"] = "Ensaios" if tipo_ui == "Ensaio Musical" else "Aulas MSA"
                st.rerun()
            except Exception as e:
                st.error(f"Não foi possível salvar o registro: {e}")

st.markdown("---")
st.caption("Agenda Musical • Região de Jaciara - MT")
