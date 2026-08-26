import calendar
from datetime import date
from io import BytesIO
from urllib.parse import quote

import requests
import streamlit as st
import xlsxwriter
from fpdf import FPDF

SUPABASE_URL = "https://ovnwnzqjjjtfqjodvusi.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_uBqke5HDz9U-xSKjxhzUww_-Y0qW367"

NOMES_MESES = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}
DIAS_SEMANA_PT = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
DIAS_SEMANA_CURTO = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]
FREQUENCIAS = ["Todos os Meses", "Meses Ímpares", "Meses Pares"]


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


def definir_notificacao(mensagem):
    st.session_state["flash_message"] = mensagem


def mostrar_notificacao():
    mensagem = st.session_state.get("flash_message")
    if mensagem:
        st.success(mensagem)
        try:
            st.toast(mensagem, icon="✅")
        except Exception:
            pass
        st.session_state["flash_message"] = None


def carregar_eventos():
    try:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/calendario_eventos",
            params={"select": "id,nome,local,dia_sem,semana,hora,interc", "order": "id.asc"},
            headers=_headers(SUPABASE_PUBLISHABLE_KEY),
            timeout=10,
        )
        r.raise_for_status()
        return r.json(), None
    except Exception as exc:
        return [], str(exc)


def carregar_avisos():
    try:
        r = requests.get(
            f"{SUPABASE_URL}/rest/v1/calendario_avisos",
            params={"select": "mes,texto", "order": "mes.asc"},
            headers=_headers(SUPABASE_PUBLISHABLE_KEY),
            timeout=10,
        )
        r.raise_for_status()
        return {int(x["mes"]): x.get("texto", "") for x in r.json()}, None
    except Exception as exc:
        return {}, str(exc)


def inserir_evento(evt):
    key = _admin_secret()
    if not key:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/calendario_eventos",
        json=evt,
        headers=_headers(key, "return=representation"),
        timeout=10,
    )
    r.raise_for_status()
    return r.json()


def atualizar_evento(evento_id, evt):
    key = _admin_secret()
    if not key:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r = requests.patch(
        f"{SUPABASE_URL}/rest/v1/calendario_eventos",
        params={"id": f"eq.{int(evento_id)}"},
        json=evt,
        headers=_headers(key, "return=representation"),
        timeout=10,
    )
    r.raise_for_status()
    dados = r.json()
    if not dados:
        raise RuntimeError("O registro não foi encontrado para atualização.")
    return dados


def excluir_evento(evento_id):
    key = _admin_secret()
    if not key:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r = requests.delete(
        f"{SUPABASE_URL}/rest/v1/calendario_eventos",
        params={"id": f"eq.{int(evento_id)}"},
        headers=_headers(key, "return=representation"),
        timeout=10,
    )
    r.raise_for_status()
    return r.json()


def salvar_aviso(mes, texto):
    key = _admin_secret()
    if not key:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r = requests.post(
        f"{SUPABASE_URL}/rest/v1/calendario_avisos",
        params={"on_conflict": "mes"},
        json={"mes": int(mes), "texto": texto},
        headers=_headers(key, "resolution=merge-duplicates,return=representation"),
        timeout=10,
    )
    r.raise_for_status()


def excluir_aviso(mes):
    key = _admin_secret()
    if not key:
        raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r = requests.delete(
        f"{SUPABASE_URL}/rest/v1/calendario_avisos",
        params={"mes": f"eq.{int(mes)}"},
        headers=_headers(key),
        timeout=10,
    )
    r.raise_for_status()


def calcular_eventos(ano, eventos):
    agenda = {}
    calendar.setfirstweekday(calendar.SUNDAY)

    for mes in range(1, 13):
        matriz = calendar.monthcalendar(ano, mes)
        for evt in eventos:
            freq = evt["interc"]
            deve_marcar = (
                freq == "Todos os Meses"
                or (freq == "Meses Ímpares" and mes % 2 != 0)
                or (freq == "Meses Pares" and mes % 2 == 0)
            )
            if not deve_marcar:
                continue

            contador = 0
            achado = None
            for semana in matriz:
                numero = semana[int(evt["dia_sem"])]
                if numero:
                    contador += 1
                    if contador == int(evt["semana"]):
                        achado = numero
                        break

            if achado:
                agenda.setdefault(f"{ano}-{mes}-{achado}", []).append({
                    "titulo": evt["nome"],
                    "local": evt["local"],
                    "hora": evt["hora"],
                })
    return agenda


def montar_agenda_ordenada(ano, eventos):
    resultado = []
    for chave, evts in calcular_eventos(ano, eventos).items():
        y, m, d = map(int, chave.split("-"))
        dt = date(y, m, d)
        resultado.extend((dt, evt) for evt in evts)
    return sorted(resultado, key=lambda x: x[0])


def gerar_link_google(dt, evt):
    hora = evt["hora"].replace("HRS", "").replace(":", "").strip()
    if len(hora) != 4 or not hora.isdigit():
        hora = "1930"

    hi = int(hora[:2])
    hf = min(hi + 2, 23)
    inicio = f"{dt.year}{dt.month:02d}{dt.day:02d}T{hi:02d}{hora[2:]}00"
    fim = f"{dt.year}{dt.month:02d}{dt.day:02d}T{hf:02d}{hora[2:]}00"

    return (
        "https://calendar.google.com/calendar/render?action=TEMPLATE"
        f"&text={quote(evt['titulo'] + ' - ' + evt['local'])}"
        f"&dates={inicio}/{fim}"
        f"&location={quote(evt['local'])}"
        "&details=Ensaio+CCB&sf=true&output=xml"
    )


def gerar_excel(ano, eventos, avisos):
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("Calendário")

    head = wb.add_format({
        "bold": True, "bg_color": "#1F4E5F", "font_color": "white",
        "align": "center", "border": 1
    })
    cell = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
    evt_fmt = wb.add_format({
        "border": 1, "text_wrap": True, "valign": "top",
        "bg_color": "#FFFF00", "bold": True
    })

    agenda = montar_agenda_ordenada(ano, eventos)
    mapa = {}
    for dt, evt in agenda:
        mapa.setdefault((dt.month, dt.day), []).append(evt)

    ws.set_column(0, 6, 18)
    row = 0

    for mes in range(1, 13):
        ws.merge_range(row, 0, row, 6, f"{NOMES_MESES[mes]} {ano}", head)
        row += 1

        for col, dia in enumerate(DIAS_SEMANA_CURTO):
            ws.write(row, col, dia, head)
        row += 1

        for semana in calendar.monthcalendar(ano, mes):
            ws.set_row(row, 65)
            for col, dia in enumerate(semana):
                texto = "" if not dia else str(dia)
                fmt = cell

                if dia and (mes, dia) in mapa:
                    texto += "\n" + "\n".join(
                        f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}"
                        for evt in mapa[(mes, dia)]
                    )
                    fmt = evt_fmt

                ws.write(row, col, texto, fmt)
            row += 1

        ws.merge_range(row, 0, row, 6, f"Anotações: {avisos.get(mes, '')}", cell)
        row += 2

    wb.close()
    out.seek(0)
    return out


def texto_pdf_seguro(valor):
    texto = str(valor or "")
    texto = texto.replace("—", "-").replace("–", "-").replace("•", "-")
    return texto.encode("latin-1", "replace").decode("latin-1")


def gerar_pdf(ano, eventos, avisos):
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)

    agenda = montar_agenda_ordenada(ano, eventos)
    eventos_dict = {}
    for dt, evt in agenda:
        eventos_dict.setdefault((dt.month, dt.day), []).append(evt)

    margem = 10
    largura_total = 190
    largura_coluna = largura_total / 7
    altura_cabecalho = 8

    for mes in range(1, 13):
        pdf.add_page()

        # Faixa do mês igual ao layout original
        pdf.set_fill_color(31, 78, 95)
        pdf.rect(margem, 10, largura_total, 15, "F")
        pdf.set_xy(margem, 10)
        pdf.set_font("Arial", "B", 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(largura_total, 15, texto_pdf_seguro(f"{NOMES_MESES[mes]} {ano}"), 0, 0, "C")

        # Cabeçalho dos dias da semana
        y_header = 30
        pdf.set_xy(margem, y_header)
        pdf.set_font("Arial", "B", 7)
        pdf.set_fill_color(31, 78, 95)
        pdf.set_text_color(255, 255, 255)
        for dia_semana in DIAS_SEMANA_CURTO:
            pdf.cell(largura_coluna, altura_cabecalho, dia_semana, 1, 0, "C", fill=True)

        matriz = calendar.monthcalendar(ano, mes)
        numero_semanas = len(matriz)
        topo_grade = y_header + altura_cabecalho
        espaco_avisos = 28
        fundo_grade = 260 - topo_grade - espaco_avisos
        altura_linha = fundo_grade / numero_semanas

        for linha_idx, semana in enumerate(matriz):
            y = topo_grade + linha_idx * altura_linha

            for col_idx, dia in enumerate(semana):
                x = margem + col_idx * largura_coluna
                tem_evento = dia != 0 and (mes, dia) in eventos_dict

                if dia == 0:
                    pdf.set_fill_color(230, 230, 230)
                elif tem_evento:
                    pdf.set_fill_color(255, 255, 0)
                else:
                    pdf.set_fill_color(255, 255, 255)

                pdf.set_draw_color(0, 0, 0)
                pdf.rect(x, y, largura_coluna, altura_linha, "DF")

                if dia == 0:
                    continue

                # Número do dia
                pdf.set_xy(x + 1.5, y + 1.2)
                pdf.set_font("Arial", "B", 8)
                pdf.set_text_color(0, 0, 0)
                pdf.cell(6, 4, str(dia), 0, 0, "L")

                if tem_evento:
                    cursor_y = y + 6
                    for evt in eventos_dict[(mes, dia)]:
                        texto_evt = texto_pdf_seguro(
                            f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}"
                        )

                        # Caixa interna fixa: evita o erro de multi_cell sem espaço horizontal
                        pdf.set_xy(x + 1.5, cursor_y)
                        pdf.set_font("Arial", "B", 5.3)
                        pdf.set_text_color(0, 0, 0)
                        pdf.multi_cell(
                            largura_coluna - 3,
                            2.5,
                            texto_evt,
                            border=0,
                            align="L",
                        )
                        cursor_y = pdf.get_y() + 1

                        # Não deixa texto extrapolar a célula
                        if cursor_y > y + altura_linha - 2:
                            break

        # Área de anotações / avisos
        y_aviso = topo_grade + numero_semanas * altura_linha + 4
        aviso = texto_pdf_seguro(avisos.get(mes, ""))
        pdf.set_xy(margem, y_aviso)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 9)

        if aviso:
            pdf.set_fill_color(255, 230, 230)
            pdf.cell(largura_total, 6, "Anotacoes / Avisos Importantes:", 1, 1, "L", fill=True)
            pdf.set_x(margem)
            pdf.set_font("Arial", "B", 9)
            pdf.set_text_color(180, 0, 0)
            pdf.multi_cell(largura_total, 6, aviso, border=1, align="L", fill=True)
        else:
            pdf.set_fill_color(255, 255, 255)
            pdf.cell(largura_total, 6, "Anotacoes:", 1, 1, "L", fill=True)
            pdf.set_x(margem)
            pdf.cell(largura_total, 10, "", 1, 0, "L", fill=True)

    valor = pdf.output(dest="S")
    return valor.encode("latin-1") if isinstance(valor, str) else bytes(valor)


st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered")

for chave, valor in {
    "theme": "light",
    "nav": "Agenda",
    "ano_base": date.today().year,
    "flash_message": None,
}.items():
    if chave not in st.session_state:
        st.session_state[chave] = valor


dark = st.session_state.theme == "dark"
bg = "linear-gradient(135deg,#0F2027,#203A43,#2C5364)" if dark else "linear-gradient(135deg,#F5F7FA,#C3CFE2)"
text = "#fff" if dark else "#1F4E5F"
card = "rgba(30,40,50,.78)" if dark else "rgba(255,255,255,.9)"

st.markdown(f"""
<style>
#MainMenu, footer, header {{visibility:hidden}}
.stApp {{background:{bg}; background-attachment:fixed}}
.block-container {{max-width:780px; padding-top:1.5rem}}
.title {{text-align:center; color:{text}}}
.event-card {{background:{card}; border-radius:18px; padding:16px; margin:12px 0; box-shadow:0 8px 28px rgba(0,0,0,.12)}}
.month {{font-size:26px; font-weight:900; color:{text}; margin:30px 0 8px}}
.next {{background:linear-gradient(135deg,#1F4E5F,#468196); color:white; padding:20px; border-radius:20px; text-align:center; margin-bottom:24px}}
.aviso {{background:#fff0f0; color:#b71c1c; border-left:4px solid #d32f2f; padding:12px; border-radius:8px}}
</style>
""", unsafe_allow_html=True)

_, theme_col = st.columns([8, 1])
with theme_col:
    if st.button("☀️" if dark else "🌙"):
        st.session_state.theme = "light" if dark else "dark"
        st.rerun()

st.markdown("<h1 class='title'>Agenda CCB Jaciara</h1>", unsafe_allow_html=True)

nav1, nav2 = st.columns(2)
if nav1.button("📅 VER AGENDA", use_container_width=True):
    st.session_state.nav = "Agenda"
    st.rerun()
if nav2.button("🔒 ADMIN", use_container_width=True):
    st.session_state.nav = "Admin"
    st.rerun()

st.divider()

eventos, erro_eventos = carregar_eventos()
avisos, erro_avisos = carregar_avisos()

if erro_eventos or erro_avisos:
    st.error("Não foi possível consultar o banco de dados agora.")

if st.session_state.nav == "Agenda":
    agenda = montar_agenda_ordenada(st.session_state.ano_base, eventos)
    hoje = date.today()
    prox = next(((dt, evt) for dt, evt in agenda if dt >= hoje), None)

    if prox:
        dt, evt = prox
        faltam = (dt - hoje).days
        txt = "HOJE!" if faltam == 0 else f"Faltam {faltam} dias"
        st.markdown(
            f"<div class='next'>✨ PRÓXIMO ENSAIO • {txt}"
            f"<h2>{evt['titulo']}</h2>{evt['local']}<br>"
            f"{dt:%d/%m/%Y} • {evt['hora']}</div>",
            unsafe_allow_html=True,
        )

    mes_atual = 0
    for dt, evt in agenda:
        if dt.month != mes_atual:
            mes_atual = dt.month
            st.markdown(
                f"<div class='month'>{NOMES_MESES[mes_atual]} {dt.year}</div>",
                unsafe_allow_html=True,
            )
            if avisos.get(mes_atual):
                st.markdown(
                    f"<div class='aviso'>📢 {avisos[mes_atual]}</div>",
                    unsafe_allow_html=True,
                )

        st.markdown(
            f"<div class='event-card'><b>{dt:%d/%m} • {evt['titulo']}</b><br>"
            f"📍 {evt['local']}<br>"
            f"🕒 {evt['hora']} • {DIAS_SEMANA_PT[int(dt.strftime('%w'))]}<br>"
            f"<a href='{gerar_link_google(dt, evt)}' target='_blank'>🔔 Adicionar lembrete</a></div>",
            unsafe_allow_html=True,
        )

else:
    st.subheader("🔒 Painel Administrativo")

    if not _admin_password() or not _admin_secret():
        st.warning("Configure ADMIN_PASSWORD e SUPABASE_SECRET_KEY nos Secrets do Streamlit.")
    else:
        senha = st.text_input("Senha de Acesso", type="password")

        if senha == _admin_password():
            st.success("✅ Acesso liberado")
            mostrar_notificacao()

            st.session_state.ano_base = st.number_input(
                "Ano de Referência",
                min_value=2020,
                max_value=2100,
                value=int(st.session_state.ano_base),
                step=1,
            )

            abas = st.tabs([
                "➕ Novo Evento",
                "📝 Avisos",
                "✏️ Gerenciar Eventos",
                "📥 Downloads",
            ])

            with abas[0]:
                with st.form("novo", clear_on_submit=True):
                    nome = st.text_input("Nome", "ENSAIO LOCAL")
                    local = st.text_input("Local")
                    dia = st.selectbox("Dia", range(7), format_func=lambda x: DIAS_SEMANA_PT[x])
                    semana = st.selectbox("Semana", [1, 2, 3, 4, 5])
                    hora = st.text_input("Hora", "19:30 HRS")
                    freq = st.selectbox("Frequência", FREQUENCIAS)

                    if st.form_submit_button("💾 Salvar Evento", use_container_width=True):
                        if not local.strip():
                            st.error("Informe o local.")
                        else:
                            try:
                                inserir_evento({
                                    "nome": nome.strip().upper(),
                                    "local": local.strip().upper(),
                                    "dia_sem": dia,
                                    "semana": semana,
                                    "hora": hora.strip().upper(),
                                    "interc": freq,
                                })
                                definir_notificacao("✅ Evento salvo com sucesso!")
                                st.rerun()
                            except Exception as exc:
                                st.error(f"Erro ao salvar: {exc}")

            with abas[1]:
                mes = st.selectbox("Mês", range(1, 13), format_func=lambda x: NOMES_MESES[x])
                aviso = st.text_area("Aviso", value=avisos.get(mes, ""))
                col_salvar, col_apagar = st.columns(2)

                if col_salvar.button("💾 Salvar Aviso", use_container_width=True):
                    try:
                        salvar_aviso(mes, aviso)
                        definir_notificacao("✅ Aviso salvo com sucesso!")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro: {exc}")

                if col_apagar.button("🗑️ Apagar Aviso", use_container_width=True):
                    try:
                        excluir_aviso(mes)
                        definir_notificacao("✅ Aviso excluído com sucesso!")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro: {exc}")

            with abas[2]:
                if not eventos:
                    st.info("Nenhum evento cadastrado.")
                else:
                    opcoes = {
                        evt["id"]: f"{evt['local']} — {evt['semana']}ª {DIAS_SEMANA_CURTO[int(evt['dia_sem'])]} — {evt['hora']}"
                        for evt in eventos
                    }

                    evento_id = st.selectbox(
                        "Selecione o evento",
                        list(opcoes),
                        format_func=lambda x: opcoes[x],
                    )
                    evt = next(e for e in eventos if e["id"] == evento_id)

                    with st.form(f"editar_{evento_id}"):
                        enome = st.text_input("Nome", evt["nome"])
                        elocal = st.text_input("Local", evt["local"])
                        edia = st.selectbox(
                            "Dia",
                            range(7),
                            index=int(evt["dia_sem"]),
                            format_func=lambda x: DIAS_SEMANA_PT[x],
                        )
                        esemana = st.selectbox(
                            "Semana",
                            [1, 2, 3, 4, 5],
                            index=int(evt["semana"]) - 1,
                        )
                        ehora = st.text_input("Hora", evt["hora"])
                        efreq = st.selectbox(
                            "Frequência",
                            FREQUENCIAS,
                            index=FREQUENCIAS.index(evt["interc"]),
                        )

                        if st.form_submit_button("✅ SALVAR ALTERAÇÕES", use_container_width=True):
                            if not elocal.strip():
                                st.error("Informe o local.")
                            else:
                                try:
                                    atualizar_evento(evento_id, {
                                        "nome": enome.strip().upper(),
                                        "local": elocal.strip().upper(),
                                        "dia_sem": edia,
                                        "semana": esemana,
                                        "hora": ehora.strip().upper(),
                                        "interc": efreq,
                                    })
                                    definir_notificacao("✅ Alterações salvas com sucesso!")
                                    st.rerun()
                                except Exception as exc:
                                    st.error(f"Erro ao atualizar: {exc}")

                    st.divider()
                    st.warning(f"Excluir permanentemente: {evt['nome']} — {evt['local']}")
                    confirmar = st.checkbox(
                        "Confirmo que desejo excluir este evento",
                        key=f"confirmar_exclusao_{evento_id}",
                    )

                    if st.button(
                        "🗑️ EXCLUIR EVENTO",
                        use_container_width=True,
                        disabled=not confirmar,
                        type="secondary",
                        key=f"excluir_{evento_id}",
                    ):
                        try:
                            excluir_evento(evento_id)
                            definir_notificacao("✅ Evento excluído com sucesso!")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Erro ao excluir: {exc}")

            with abas[3]:
                try:
                    arquivo_excel = gerar_excel(st.session_state.ano_base, eventos, avisos)
                    st.download_button(
                        "⬇️ Baixar Excel",
                        arquivo_excel,
                        f"Calendario_{st.session_state.ano_base}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"Não foi possível gerar o Excel: {exc}")

                try:
                    arquivo_pdf = gerar_pdf(st.session_state.ano_base, eventos, avisos)
                    st.download_button(
                        "⬇️ Baixar PDF",
                        arquivo_pdf,
                        f"Calendario_{st.session_state.ano_base}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception as exc:
                    st.error(f"Não foi possível gerar o PDF: {exc}")

        elif senha:
            st.error("❌ Senha incorreta")
