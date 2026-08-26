import calendar
from datetime import date
from io import BytesIO
from urllib.parse import quote

import requests
import streamlit as st
import xlsxwriter
from fpdf import FPDF

# ==========================================
# CONFIGURAÇÃO SUPABASE
# ==========================================
SUPABASE_URL = "https://ovnwnzqjjjtfqjodvusi.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_uBqke5HDz9U-xSKjxhzUww_-Y0qW367"

NOMES_MESES = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}
DIAS_SEMANA_PT = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
DIAS_SEMANA_CURTO = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]


def _headers(key, prefer=None):
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def _admin_secret():
    return st.secrets.get("SUPABASE_SECRET_KEY", "")


def _admin_password():
    return st.secrets.get("ADMIN_PASSWORD", "")


def carregar_eventos():
    try:
        resp = requests.get(
            f"{SUPABASE_URL}/rest/v1/calendario_eventos",
            params={"select": "id,nome,local,dia_sem,semana,hora,interc", "order": "id.asc"},
            headers=_headers(SUPABASE_PUBLISHABLE_KEY),
            timeout=10,
        )
        resp.raise_for_status()
        return resp.json(), None
    except Exception as exc:
        return [], str(exc)


def carregar_avisos():
    try:
        resp = requests.get(
            f"{SUPABASE_URL}/rest/v1/calendario_avisos",
            params={"select": "mes,texto", "order": "mes.asc"},
            headers=_headers(SUPABASE_PUBLISHABLE_KEY),
            timeout=10,
        )
        resp.raise_for_status()
        return {int(item["mes"]): item.get("texto", "") for item in resp.json()}, None
    except Exception as exc:
        return {}, str(exc)


def inserir_evento(evento):
    secret = _admin_secret()
    if not secret:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada nos Secrets do Streamlit.")
    resp = requests.post(
        f"{SUPABASE_URL}/rest/v1/calendario_eventos",
        json=evento,
        headers=_headers(secret, "return=representation"),
        timeout=10,
    )
    resp.raise_for_status()
    return resp.json()


def excluir_evento(evento_id):
    secret = _admin_secret()
    if not secret:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada nos Secrets do Streamlit.")
    resp = requests.delete(
        f"{SUPABASE_URL}/rest/v1/calendario_eventos",
        params={"id": f"eq.{evento_id}"},
        headers=_headers(secret),
        timeout=10,
    )
    resp.raise_for_status()


def salvar_aviso(mes, texto):
    secret = _admin_secret()
    if not secret:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada nos Secrets do Streamlit.")
    resp = requests.post(
        f"{SUPABASE_URL}/rest/v1/calendario_avisos",
        params={"on_conflict": "mes"},
        json={"mes": int(mes), "texto": texto},
        headers=_headers(secret, "resolution=merge-duplicates,return=representation"),
        timeout=10,
    )
    resp.raise_for_status()


def excluir_aviso(mes):
    secret = _admin_secret()
    if not secret:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada nos Secrets do Streamlit.")
    resp = requests.delete(
        f"{SUPABASE_URL}/rest/v1/calendario_avisos",
        params={"mes": f"eq.{int(mes)}"},
        headers=_headers(secret),
        timeout=10,
    )
    resp.raise_for_status()


# ==========================================
# LÓGICA DO CALENDÁRIO
# ==========================================
def calcular_eventos(ano, lista_eventos):
    agenda = {}
    calendar.setfirstweekday(calendar.SUNDAY)
    for mes in range(1, 13):
        cal_matrix = calendar.monthcalendar(ano, mes)
        for evt in lista_eventos:
            interc = evt["interc"]
            deve_marcar = (
                interc == "Todos os Meses"
                or (interc == "Meses Ímpares" and mes % 2 != 0)
                or (interc == "Meses Pares" and mes % 2 == 0)
            )
            if not deve_marcar:
                continue

            contador = 0
            dia_encontrado = None
            dia_alvo_idx = int(evt["dia_sem"])
            semana_alvo = int(evt["semana"])
            for semana in cal_matrix:
                dia_num = semana[dia_alvo_idx]
                if dia_num != 0:
                    contador += 1
                    if contador == semana_alvo:
                        dia_encontrado = dia_num
                        break

            if dia_encontrado:
                chave = f"{ano}-{mes}-{dia_encontrado}"
                agenda.setdefault(chave, []).append({
                    "titulo": evt["nome"],
                    "local": evt["local"],
                    "hora": evt["hora"],
                })
    return agenda


def montar_agenda_ordenada(ano, lista_eventos):
    dados = calcular_eventos(ano, lista_eventos)
    lista_final = []
    for chave, eventos in dados.items():
        y, m, d = map(int, chave.split("-"))
        dt = date(y, m, d)
        for evt in eventos:
            lista_final.append((dt, evt))
    lista_final.sort(key=lambda x: x[0])
    return lista_final


def gerar_link_google(dt, evt_data):
    titulo = quote(f"{evt_data['titulo']} - {evt_data['local']}")
    hora_limpa = evt_data["hora"].replace("HRS", "").replace(":", "").strip()
    if len(hora_limpa) != 4 or not hora_limpa.isdigit():
        hora_limpa = "1930"
    hora = int(hora_limpa[:2])
    minuto = hora_limpa[2:]
    fim_hora = min(hora + 2, 23)
    data_inicio = f"{dt.year}{dt.month:02d}{dt.day:02d}T{hora:02d}{minuto}00"
    data_fim = f"{dt.year}{dt.month:02d}{dt.day:02d}T{fim_hora:02d}{minuto}00"
    local = quote(evt_data["local"])
    return (
        "https://calendar.google.com/calendar/render?action=TEMPLATE"
        f"&text={titulo}&dates={data_inicio}/{data_fim}&location={local}"
        "&details=Ensaio+CCB&sf=true&output=xml"
    )


# ==========================================
# EXPORTAÇÃO EXCEL / PDF
# ==========================================
def gerar_excel_todos_meses(ano, lista_eventos, avisos):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("Calendário")
    header_mes = wb.add_format({"bold": True, "font_size": 14, "bg_color": "#1F4E5F", "font_color": "white", "align": "center", "valign": "vcenter", "border": 1})
    header_dias = wb.add_format({"bold": True, "bg_color": "#1F4E5F", "font_color": "white", "align": "center", "valign": "vcenter", "border": 1})
    cell_dia = wb.add_format({"border": 1, "align": "left", "valign": "top", "bold": True})
    cell_evento = wb.add_format({"border": 1, "align": "left", "valign": "top", "font_size": 8, "text_wrap": True, "bg_color": "#FFFF00", "bold": True})
    cell_aviso = wb.add_format({"border": 1, "align": "left", "text_wrap": True, "bg_color": "#FFCDD2", "bold": True, "font_color": "#B71C1C"})
    cell_vazio = wb.add_format({"border": 1, "bg_color": "#E0E0E0"})

    agenda = montar_agenda_ordenada(ano, lista_eventos)
    eventos_dict = {}
    for dt, evt in agenda:
        eventos_dict.setdefault(f"{dt.year}-{dt.month}-{dt.day}", []).append(evt)

    for col in range(7):
        ws.set_column(col, col, 18)

    row = 0
    for mes in range(1, 13):
        ws.merge_range(row, 0, row, 6, f"{NOMES_MESES[mes]} {ano}", header_mes)
        row += 1
        for col, dia in enumerate(DIAS_SEMANA_CURTO):
            ws.write(row, col, dia, header_dias)
        row += 1

        for semana in calendar.monthcalendar(ano, mes):
            ws.set_row(row, 70)
            for col, dia in enumerate(semana):
                if dia == 0:
                    ws.write(row, col, "", cell_vazio)
                    continue
                chave = f"{ano}-{mes}-{dia}"
                if chave in eventos_dict:
                    texto = f"{dia}\n"
                    for evt in eventos_dict[chave]:
                        texto += f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}\n"
                    ws.write(row, col, texto, cell_evento)
                else:
                    ws.write(row, col, dia, cell_dia)
            row += 1

        aviso = avisos.get(mes, "")
        row += 1
        ws.merge_range(row, 0, row, 6, f"Anotações: {aviso}", cell_aviso if aviso else cell_dia)
        row += 2

    wb.close()
    output.seek(0)
    return output


def gerar_pdf_calendario(ano, lista_eventos, avisos):
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    eventos_dict = {}
    for dt, evt in agenda:
        eventos_dict.setdefault(f"{dt.year}-{dt.month}-{dt.day}", []).append(evt)

    for mes in range(1, 13):
        pdf.add_page()
        pdf.set_fill_color(31, 78, 95)
        pdf.rect(10, 10, 190, 15, "F")
        pdf.set_xy(10, 10)
        pdf.set_font("Arial", "B", 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(190, 15, f"{NOMES_MESES[mes]} {ano}", 0, 1, "C")

        col_width = 27.1
        row_height = 30
        pdf.set_xy(10, 30)
        pdf.set_font("Arial", "B", 8)
        for dia in DIAS_SEMANA_CURTO:
            pdf.cell(col_width, 8, dia, 1, 0, "C", fill=True)

        y = 38
        for semana in calendar.monthcalendar(ano, mes):
            x = 10
            for dia in semana:
                chave = f"{ano}-{mes}-{dia}"
                if dia == 0:
                    pdf.set_fill_color(230, 230, 230)
                elif chave in eventos_dict:
                    pdf.set_fill_color(255, 255, 0)
                else:
                    pdf.set_fill_color(255, 255, 255)
                pdf.set_xy(x, y)
                pdf.cell(col_width, row_height, "", 1, 0, "C", fill=True)
                if dia:
                    pdf.set_xy(x + 1, y + 1)
                    pdf.set_text_color(0, 0, 0)
                    pdf.set_font("Arial", "B", 10)
                    pdf.cell(5, 5, str(dia))
                    if chave in eventos_dict:
                        pdf.set_xy(x + 1, y + 6)
                        pdf.set_font("Arial", "B", 6)
                        texto = ""
                        for evt in eventos_dict[chave]:
                            texto += f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}\n"
                        pdf.multi_cell(col_width - 2, 3, texto)
                x += col_width
            y += row_height

        aviso = avisos.get(mes, "")
        pdf.set_xy(10, 260)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 10)
        pdf.cell(190, 6, "Anotacoes / Avisos:", "LTR", 1, "L")
        pdf.set_font("Arial", "", 9)
        pdf.multi_cell(190, 15, aviso, "LBR", "L")

    val = pdf.output(dest="S")
    return val.encode("latin-1") if isinstance(val, str) else bytes(val)


# ==========================================
# INTERFACE
# ==========================================
st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered", initial_sidebar_state="collapsed")

if "theme" not in st.session_state:
    st.session_state.theme = "light"
if "nav" not in st.session_state:
    st.session_state.nav = "Agenda"
if "ano_base" not in st.session_state:
    st.session_state.ano_base = date.today().year

is_dark = st.session_state.theme == "dark"
bg = "linear-gradient(135deg,#0F2027,#203A43,#2C5364)" if is_dark else "linear-gradient(135deg,#F5F7FA,#C3CFE2)"
text = "#FFFFFF" if is_dark else "#1F4E5F"
card = "rgba(30,40,50,.78)" if is_dark else "rgba(255,255,255,.88)"
secondary = "#CFD8DC" if is_dark else "#546E7A"

st.markdown(f"""
<style>
#MainMenu, footer, header {{visibility:hidden;}}
.stApp {{background:{bg}; background-attachment:fixed;}}
.block-container {{max-width:780px; padding-top:1.5rem; padding-bottom:4rem;}}
.title {{text-align:center; color:{text}; margin-bottom:4px;}}
.subtitle {{text-align:center; color:{secondary}; margin-bottom:22px;}}
.event-card {{background:{card}; border-radius:18px; padding:16px; margin:12px 0; box-shadow:0 8px 28px rgba(0,0,0,.12);}}
.event-title {{font-size:18px; font-weight:800; color:{text};}}
.event-info {{color:{secondary}; margin-top:4px;}}
.month {{font-size:26px; font-weight:900; color:{text}; margin:32px 0 10px;}}
.next {{background:linear-gradient(135deg,#1F4E5F,#468196); color:white; padding:20px; border-radius:20px; text-align:center; margin-bottom:24px;}}
.aviso {{background:#fff0f0; color:#b71c1c; border-left:4px solid #d32f2f; padding:12px; border-radius:8px; margin-bottom:12px;}}
</style>
""", unsafe_allow_html=True)

c1, c2 = st.columns([8, 1])
with c2:
    if st.button("☀️" if is_dark else "🌙"):
        st.session_state.theme = "light" if is_dark else "dark"
        st.rerun()

st.markdown("<h1 class='title'>Agenda CCB Jaciara</h1><div class='subtitle'>Consulte datas e horários oficiais</div>", unsafe_allow_html=True)

nav1, nav2 = st.columns(2)
with nav1:
    if st.button("📅 VER AGENDA", use_container_width=True):
        st.session_state.nav = "Agenda"
        st.rerun()
with nav2:
    if st.button("🔒 ADMIN", use_container_width=True):
        st.session_state.nav = "Admin"
        st.rerun()

st.divider()

eventos, erro_eventos = carregar_eventos()
avisos, erro_avisos = carregar_avisos()

if erro_eventos or erro_avisos:
    st.error("Não foi possível consultar o banco de dados agora.")
    if erro_eventos:
        st.caption(f"Eventos: {erro_eventos}")
    if erro_avisos:
        st.caption(f"Avisos: {erro_avisos}")

if st.session_state.nav == "Agenda":
    agenda = montar_agenda_ordenada(st.session_state.ano_base, eventos)
    hoje = date.today()
    prox = next(((dt, evt) for dt, evt in agenda if dt >= hoje), None)

    if prox:
        dt, evt = prox
        faltam = (dt - hoje).days
        txt = "HOJE!" if faltam == 0 else f"Faltam {faltam} dias"
        st.markdown(
            f"<div class='next'><small>✨ PRÓXIMO ENSAIO • {txt}</small>"
            f"<h2>{evt['titulo']}</h2><div>{evt['local']}</div>"
            f"<div>{dt.day:02d}/{dt.month:02d}/{dt.year} • {DIAS_SEMANA_PT[int(dt.strftime('%w'))]} • {evt['hora']}</div></div>",
            unsafe_allow_html=True,
        )

    if not agenda:
        st.info("Nenhum evento encontrado para este ano.")
    else:
        mes_atual = 0
        for dt, evt in agenda:
            if dt.month != mes_atual:
                mes_atual = dt.month
                st.markdown(f"<div class='month'>{NOMES_MESES[mes_atual]} {dt.year}</div>", unsafe_allow_html=True)
                if avisos.get(mes_atual):
                    st.markdown(f"<div class='aviso'>📢 {avisos[mes_atual]}</div>", unsafe_allow_html=True)

            link = gerar_link_google(dt, evt)
            st.markdown(
                f"<div class='event-card'><div class='event-title'>{dt.day:02d}/{dt.month:02d} • {evt['titulo']}</div>"
                f"<div class='event-info'>📍 {evt['local']}</div>"
                f"<div class='event-info'>🕒 {evt['hora']} • {DIAS_SEMANA_PT[int(dt.strftime('%w'))]}</div>"
                f"<div style='margin-top:10px'><a href='{link}' target='_blank'>🔔 Adicionar lembrete</a></div></div>",
                unsafe_allow_html=True,
            )

else:
    st.subheader("🔒 Painel Administrativo")

    admin_password = _admin_password()
    admin_secret = _admin_secret()

    if not admin_password or not admin_secret:
        st.warning("O ADMIN precisa dos Secrets do Streamlit para gravar no Supabase.")
        st.code('ADMIN_PASSWORD = "sua_senha"\nSUPABASE_SECRET_KEY = "sb_secret_..."', language="toml")
    else:
        senha = st.text_input("Senha de Acesso", type="password")
        if senha == admin_password:
            st.success("✅ Acesso liberado")
            st.session_state.ano_base = st.number_input("Ano de Referência", min_value=2020, max_value=2100, value=int(st.session_state.ano_base), step=1)

            abas = st.tabs(["➕ Novo Evento", "📝 Avisos", "📋 Gerenciar", "📥 Downloads"])

            with abas[0]:
                with st.form("novo_evento", clear_on_submit=True):
                    nome = st.text_input("Nome", "ENSAIO LOCAL")
                    local = st.text_input("Local")
                    dia = st.selectbox("Dia", range(7), format_func=lambda x: DIAS_SEMANA_PT[x])
                    semana = st.selectbox("Semana", [1, 2, 3, 4, 5])
                    hora = st.text_input("Hora", "19:30 HRS")
                    interc = st.selectbox("Frequência", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
                    if st.form_submit_button("💾 Salvar Evento", use_container_width=True):
                        if not local.strip():
                            st.error("Informe o local.")
                        else:
                            try:
                                inserir_evento({
                                    "nome": nome.strip().upper(),
                                    "local": local.strip().upper(),
                                    "dia_sem": int(dia),
                                    "semana": int(semana),
                                    "hora": hora.strip().upper(),
                                    "interc": interc,
                                })
                                st.success("Evento salvo no Supabase.")
                                st.rerun()
                            except Exception as exc:
                                st.error(f"Erro ao salvar: {exc}")

            with abas[1]:
                mes = st.selectbox("Escolha o mês", range(1, 13), format_func=lambda x: NOMES_MESES[x])
                texto = st.text_area("Texto do aviso", value=avisos.get(mes, ""), height=110)
                a1, a2 = st.columns(2)
                if a1.button("💾 Salvar Aviso", use_container_width=True):
                    try:
                        salvar_aviso(mes, texto)
                        st.success("Aviso salvo no Supabase.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro ao salvar: {exc}")
                if a2.button("🗑️ Apagar Aviso", use_container_width=True):
                    try:
                        excluir_aviso(mes)
                        st.success("Aviso apagado.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro ao apagar: {exc}")

            with abas[2]:
                if not eventos:
                    st.info("Nenhum evento cadastrado.")
                for evt in eventos:
                    m1, m2 = st.columns([5, 1])
                    m1.write(f"**{evt['local']}** — {evt['semana']}ª {DIAS_SEMANA_CURTO[int(evt['dia_sem'])]} — {evt['hora']}")
                    if m2.button("🗑️", key=f"del_{evt['id']}"):
                        try:
                            excluir_evento(evt["id"])
                            st.success("Evento excluído.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Erro ao excluir: {exc}")

            with abas[3]:
                excel = gerar_excel_todos_meses(st.session_state.ano_base, eventos, avisos)
                st.download_button(
                    "⬇️ Baixar Excel",
                    excel,
                    f"Calendario_{st.session_state.ano_base}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                pdf = gerar_pdf_calendario(st.session_state.ano_base, eventos, avisos)
                st.download_button(
                    "⬇️ Baixar PDF",
                    pdf,
                    f"Calendario_{st.session_state.ano_base}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        elif senha:
            st.error("❌ Senha incorreta")
