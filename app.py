import streamlit as st
import xlsxwriter
import calendar
from io import BytesIO
from datetime import datetime, date
from fpdf import FPDF
from urllib.parse import quote

# ==========================================
# 1. LÓGICA E FUNÇÕES (MANTIDAS)
# ==========================================
NOMES_MESES = {1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"}
DIAS_SEMANA_PT = ["DOMINGO", "SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
DIAS_SEMANA_CURTO = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SAB"]

def calcular_eventos(ano, lista_eventos):
    agenda = {}
    calendar.setfirstweekday(calendar.SUNDAY) 
    for mes in range(1, 13):
        cal_matrix = calendar.monthcalendar(ano, mes)
        for evt in lista_eventos:
            deve_marcar = False
            interc = evt["interc"]
            if interc == "Todos os Meses": deve_marcar = True
            elif interc == "Meses Ímpares" and (mes % 2 != 0): deve_marcar = True
            elif interc == "Meses Pares" and (mes % 2 == 0): deve_marcar = True

            if deve_marcar:
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
                    evento_dados = {"titulo": evt['nome'], "local": evt['local'], "hora": evt['hora']}
                    if chave not in agenda: agenda[chave] = []
                    agenda[chave].append(evento_dados)
    return agenda

def montar_agenda_ordenada(ano, lista_eventos):
    dados = calcular_eventos(ano, lista_eventos)
    lista_final = []
    for chave, lista_evts_dados in dados.items():
        parts = chave.split('-')
        dt = date(int(parts[0]), int(parts[1]), int(parts[2]))
        for evt_data in lista_evts_dados:
            lista_final.append((dt, evt_data))
    lista_final.sort(key=lambda x: x[0])
    return lista_final

def gerar_link_google(dt, evt_data):
    titulo = quote(f"{evt_data['titulo']} - {evt_data['local']}")
    hora_limpa = evt_data['hora'].replace("HRS", "").replace(":", "").strip()
    if len(hora_limpa) < 4: hora_limpa = "1930"
    data_inicio = f"{dt.year}{dt.month:02d}{dt.day:02d}T{hora_limpa}00"
    data_fim = f"{dt.year}{dt.month:02d}{dt.day:02d}T{int(hora_limpa[:2])+2:02d}{hora_limpa[2:]}00"
    local = quote(evt_data['local'])
    return f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={titulo}&dates={data_inicio}/{data_fim}&location={local}&details=Ensaio+CCB&sf=true&output=xml"

# ===== FUNÇÕES DE ARQUIVO =====
def gerar_excel_todos_meses(ano, lista_eventos, avisos):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Calendário")
    header_mes = wb.add_format({'bold': True, 'font_size': 14, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_dias = wb.add_format({'bold': True, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    cell_dia = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 'bold': True})
    cell_evento = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 8, 'text_wrap': True, 'bg_color': '#FFFF00', 'bold': True})
    cell_aviso = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 'text_wrap': True, 'bg_color': '#FFCDD2', 'bold': True, 'font_color': '#B71C1C'})
    cell_vazio = wb.add_format({'border': 1, 'bg_color': '#E0E0E0'})
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    eventos_dict = {}
    for dt, evt_data in agenda:
        chave = f"{dt.year}-{dt.month}-{dt.day}"
        if chave not in eventos_dict: eventos_dict[chave] = []
        eventos_dict[chave].append(evt_data)
    for col in range(7): ws.set_column(col, col, 18)
    current_row = 0
    for mes in range(1, 13):
        nome_mes = NOMES_MESES[mes]
        ws.merge_range(current_row, 0, current_row, 6, f"{nome_mes} {ano}", header_mes)
        ws.set_row(current_row, 30)
        current_row += 1
        for col, dia in enumerate(DIAS_SEMANA_CURTO):
            ws.write(current_row, col, dia, header_dias)
        ws.set_row(current_row, 20)
        current_row += 1
        cal_matrix = calendar.monthcalendar(ano, mes)
        for semana in cal_matrix:
            ws.set_row(current_row, 70)
            for col, dia in enumerate(semana):
                if dia == 0: ws.write(current_row, col, '', cell_vazio)
                else:
                    chave = f"{ano}-{mes}-{dia}"
                    if chave in eventos_dict:
                        texto = f"{dia}\n"
                        for evt in eventos_dict[chave]: texto += f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}\n"
                        ws.write(current_row, col, texto, cell_evento)
                    else: ws.write(current_row, col, dia, cell_dia)
            current_row += 1
        aviso_mes = avisos.get(mes, "")
        texto_anotacao = f"Anotações: {aviso_mes}"
        current_row += 1
        ws.merge_range(current_row, 0, current_row, 6, texto_anotacao, cell_aviso if aviso_mes else wb.add_format({'border': 1, 'align': 'left'}))
        current_row += 2
    wb.close()
    output.seek(0)
    return output

def gerar_pdf_calendario(ano, lista_eventos, avisos):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    eventos_dict = {}
    for dt, evt_data in agenda:
        chave = f"{dt.year}-{dt.month}-{dt.day}"
        if chave not in eventos_dict: eventos_dict[chave] = []
        eventos_dict[chave].append(evt_data)
    for mes in range(1, 13):
        pdf.add_page()
        pdf.set_fill_color(31, 78, 95)
        pdf.rect(10, 10, 190, 15, 'F')
        pdf.set_xy(10, 10)
        pdf.set_font("Arial", "B", 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(20, 15, str(ano), 0, 0, 'C')
        pdf.set_xy(30, 10)
        pdf.cell(150, 15, NOMES_MESES[mes], 0, 0, 'C')
        pdf.ln(20)
        margin_left = 10
        col_width = 27.1
        row_height = 30
        header_height = 8
        pdf.set_font("Arial", "B", 8)
        pdf.set_fill_color(31, 78, 95)
        pdf.set_text_color(255, 255, 255)
        pdf.set_x(margin_left)
        for dia in DIAS_SEMANA_CURTO: pdf.cell(col_width, header_height, dia, 1, 0, 'C', fill=True)
        pdf.ln(header_height)
        cal_matrix = calendar.monthcalendar(ano, mes)
        y_start = pdf.get_y()
        for semana in cal_matrix:
            x_current = margin_left
            for dia in semana:
                chave = f"{ano}-{mes}-{dia}"
                if dia == 0: pdf.set_fill_color(230, 230, 230)
                elif chave in eventos_dict: pdf.set_fill_color(255, 255, 0)
                else: pdf.set_fill_color(255, 255, 255)
                pdf.set_xy(x_current, y_start)
                pdf.cell(col_width, row_height, '', 1, 0, 'C', fill=True)
                if dia != 0:
                    pdf.set_xy(x_current + 1, y_start + 1)
                    pdf.set_text_color(0, 0, 0)
                    pdf.set_font("Arial", "B", 10)
                    pdf.cell(5, 5, str(dia), 0, 0)
                    if chave in eventos_dict:
                        pdf.set_xy(x_current + 1, y_start + 6)
                        pdf.set_font("Arial", "B", 6)
                        texto = ""
                        for evt in eventos_dict[chave]: texto += f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}\n"
                        pdf.multi_cell(col_width - 2, 3, texto, 0, 'L')
                x_current += col_width
            y_start += row_height
        aviso_mes = avisos.get(mes, "")
        pdf.set_xy(margin_left, 260)
        pdf.set_font("Arial", "B", 10)
        pdf.set_text_color(0, 0, 0)
        if aviso_mes:
            pdf.set_fill_color(255, 230, 230)
            pdf.cell(190, 6, "Anotacoes / Avisos Importantes:", "LTR", 1, 'L', fill=True)
            pdf.set_font("Arial", "B", 11)
            pdf.set_text_color(180, 0, 0)
            pdf.multi_cell(190, 15, aviso_mes, "LBR", 'L', fill=True)
        else:
            pdf.set_fill_color(255, 255, 255)
            pdf.cell(190, 6, "Anotacoes:", "LTR", 1, 'L')
            pdf.cell(190, 15, "", "LBR", 1, 'L')
    try:
        val = pdf.output(dest='S')
        if isinstance(val, str): return val.encode('latin-1')
        return bytes(val)
    except: return bytes(pdf.output())

# ==========================================
# 2. VISUAL
# ==========================================
st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered", initial_sidebar_state="collapsed")

if 'theme' not in st.session_state: st.session_state['theme'] = 'light'

if st.session_state['theme'] == 'light':
    css_vars = {
        'bg_gradient': 'linear-gradient(135deg, #F5F7FA 0%, #C3CFE2 100%)',
        'card_bg': 'rgba(255, 255, 255, 0.85)',
        'card_border': 'rgba(255, 255, 255, 0.6)',
        'text_color': '#1F4E5F', 'text_sec': '#546E7A',
        'shadow': '0 8px 32px 0 rgba(31, 38, 135, 0.15)',
        'highlight': '#1F4E5F', 'accent': '#FFD700',
        'menu_icon_color': '#1F4E5F',
        'admin_title_color': '#1F4E5F',
        'label_color': '#1F4E5F' # COR DOS RÓTULOS LIGHT
    }
    icon_theme = "🌙"
else:
    css_vars = {
        'bg_gradient': 'linear-gradient(135deg, #0F2027 0%, #203A43 50%, #2C5364 100%)',
        'card_bg': 'rgba(30, 40, 50, 0.75)',
        'card_border': 'rgba(255, 255, 255, 0.1)',
        'text_color': '#FFFFFF', 'text_sec': '#B0BEC5',
        'shadow': '0 8px 32px 0 rgba(0, 0, 0, 0.37)',
        'highlight': '#81D4FA', 'accent': '#FFD700',
        'menu_icon_color': '#FFFFFF',
        'admin_title_color': '#FFFFFF',
        'label_color': '#FFFFFF' # COR DOS RÓTULOS DARK
    }
    icon_theme = "☀️"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;800&display=swap');

    #MainMenu {{visibility: hidden;}} footer {{visibility: hidden;}} header {{visibility: hidden;}}

    .stApp {{
        background: {css_vars['bg_gradient']};
        background-attachment: fixed;
        font-family: 'Poppins', sans-serif;
    }}

    /* FORÇAR COR DOS LABELS DO STREAMLIT */
    .st-emotion-cache-19rxjzo, .st-emotion-cache-1qg05tj, label {{
        color: {css_vars['label_color']} !important;
        font-weight: 600 !important;
    }}
    
    /* Força cor branca ou escura nos textos gerais de markdown dentro do form */
    .stMarkdown p {{
        color: {css_vars['label_color']} !important;
    }}

    .block-container {{ padding-top: 2rem; padding-bottom: 4rem; }}

    .modern-header {{ text-align: center; padding: 20px; background: transparent; margin-bottom: 20px; }}
    .modern-header h1 {{
        font-family: 'Poppins', sans-serif; font-weight: 800; font-size: 26px;
        color: {css_vars['text_color']}; text-transform: uppercase; letter-spacing: 2px;
        text-shadow: 0px 2px 4px rgba(0,0,0,0.1); margin: 0;
    }}
    .modern-header p {{ color: {css_vars['text_sec']}; font-size: 12px; font-weight: 400; }}

    @keyframes fadeInUp {{ from {{ opacity: 0; transform: translate3d(0, 30px, 0); }} to {{ opacity: 1; transform: translate3d(0, 0, 0); }} }}

    .event-card {{
        background: {css_vars['card_bg']}; backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px);
        border-radius: 20px; border: 1px solid {css_vars['card_border']}; box-shadow: {css_vars['shadow']};
        padding: 18px; margin-bottom: 18px; display: flex; align-items: center; gap: 15px;
        animation: fadeInUp 0.6s ease-both; transition: transform 0.2s ease, box-shadow 0.2s ease;
    }}
    .event-card:hover {{ transform: translateY(-3px); box-shadow: 0 12px 40px 0 rgba(0,0,0,0.2); }}

    .date-box {{
        background: linear-gradient(135deg, #1F4E5F 0%, #14323F 100%); border-radius: 14px;
        width: 65px; height: 65px; display: flex; flex-direction: column; align-items: center;
        justify-content: center; color: white; box-shadow: 0 4px 10px rgba(31, 78, 95, 0.3);
    }}
    .date-day {{ font-size: 24px; font-weight: 700; line-height: 1; }}
    .date-month {{ font-size: 9px; font-weight: 600; text-transform: uppercase; margin-top: 2px; letter-spacing: 1px; }}

    .event-details {{ flex-grow: 1; }}
    .event-badge {{ background-color: {css_vars['accent']}; color: #333; font-size: 10px; font-weight: 800; padding: 3px 10px; border-radius: 20px; text-transform: uppercase; display: inline-block; margin-bottom: 6px; }}
    .event-title {{ font-size: 16px; font-weight: 700; color: {css_vars['text_color']}; margin: 2px 0 4px 0; }}
    .event-info {{ font-size: 13px; color: {css_vars['text_sec']}; display: flex; align-items: center; gap: 6px; margin-top: 3px; }}

    .btn-notify {{ display: inline-block; margin-top: 10px; background-color: transparent; color: {css_vars['highlight']}; font-size: 12px; font-weight: 700; padding: 6px 14px; border-radius: 20px; text-decoration: none; border: 2px solid {css_vars['highlight']}; transition: all 0.3s ease; }}
    .btn-notify:hover {{ background-color: {css_vars['highlight']}; color: white; }}

    .month-separator {{ margin: 40px 0 20px 0; display: flex; align-items: center; }}
    .month-text {{ font-size: 28px !important; font-weight: 900 !important; background: -webkit-linear-gradient(45deg, {css_vars['text_color']}, {css_vars['text_sec']}); -webkit-background-clip: text; -webkit-text-fill-color: transparent; text-transform: uppercase; letter-spacing: 2px; }}
    
    .next-event-box {{
        background: linear-gradient(135deg, #1F4E5F 0%, #468196 100%); border-radius: 20px; padding: 20px; color: white; margin-bottom: 30px; box-shadow: 0 10px 30px rgba(31, 78, 95, 0.4); text-align: center; animation: fadeInUp 0.8s ease-both;
    }}
    .next-label {{ font-size: 12px; text-transform: uppercase; letter-spacing: 2px; opacity: 0.8; }}
    .next-title {{ font-size: 22px; font-weight: 800; margin: 5px 0; }}
    .next-date {{ font-size: 16px; font-weight: 500; background: rgba(255,255,255,0.2); padding: 5px 15px; border-radius: 20px; display: inline-block; margin-top: 10px;}}

    .aviso-card {{ background: rgba(255, 0, 0, 0.05); border-left: 4px solid #D32F2F; padding: 15px; margin: 10px 0 20px 0; border-radius: 8px; color: #D32F2F; font-weight: 600; display: flex; align-items: center; gap: 10px; backdrop-filter: blur(5px); }}
    .admin-container {{ background: {css_vars['card_bg']}; padding: 25px; border-radius: 20px; box-shadow: {css_vars['shadow']}; backdrop-filter: blur(10px); }}
    
    .admin-title {{
        color: {css_vars['admin_title_color']} !important;
        font-weight: 800;
        font-size: 22px;
        margin-bottom: 15px;
    }}
</style>
""", unsafe_allow_html=True)

# BOTÃO DE TEMA FLUTUANTE
c_float_1, c_float_2 = st.columns([8, 1])
with c_float_2:
    if st.button(icon_theme, key="float_theme"):
        st.session_state['theme'] = 'dark' if st.session_state['theme'] == 'light' else 'light'
        st.rerun()

# ==========================================
# 3. DADOS E NAVEGAÇÃO
# ==========================================
if 'eventos' not in st.session_state:
    st.session_state['eventos'] = [
        {"nome": "ENSAIO LOCAL", "semana": "1", "dia_sem": "6", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "SÃO PEDRO DA CIPA - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "2", "dia_sem": "5", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "SANTA ELVIRA - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "2", "dia_sem": "6", "interc": "Todos os Meses", "hora": "17:30 HRS", "local": "SÃO LOURENÇO DE FATIMA - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "2", "dia_sem": "0", "interc": "Todos os Meses", "hora": "16:30 HRS", "local": "JARDIM BOA ESPERANÇA - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "3", "dia_sem": "1", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "CENTRAL JACIARA - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "3", "dia_sem": "2", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "JUSCIMEIRA - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "3", "dia_sem": "3", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "VILA PLANALTO - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "3", "dia_sem": "5", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "SANTO ANTONIO - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "4", "dia_sem": "6", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "DOM AQUINO - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "3", "dia_sem": "4", "interc": "Meses Ímpares", "hora": "19:30 HRS", "local": "ENTRE RIOS - MT"},
        {"nome": "ENSAIO LOCAL", "semana": "3", "dia_sem": "6", "interc": "Meses Pares", "hora": "19:30 HRS", "local": "DISTRITO DE CELMA - MT"},
    ]
if 'avisos' not in st.session_state: st.session_state['avisos'] = {} 
if 'nav' not in st.session_state: st.session_state['nav'] = 'Agenda'
if 'ano_base' not in st.session_state: st.session_state['ano_base'] = date.today().year + 1

st.markdown(f"""
<div class="modern-header">
    <h1>Agenda CCB Jaciara</h1>
    <p>Consulte datas e horários oficiais</p>
</div>
""", unsafe_allow_html=True)

# BOTÕES DE NAVEGAÇÃO
col_nav_1, col_nav_2 = st.columns(2)
with col_nav_1:
    if st.button("📅 VER AGENDA", use_container_width=True):
        st.session_state['nav'] = 'Agenda'
        st.rerun()
with col_nav_2:
    if st.button("🔒 ADMIN", use_container_width=True):
        st.session_state['nav'] = 'Admin'
        st.rerun()

st.markdown("---")

# ==========================================
# 4. PÁGINA CONTEÚDO
# ==========================================
if st.session_state['nav'] == 'Agenda':
    agenda = montar_agenda_ordenada(st.session_state['ano_base'], st.session_state['eventos'])
    
    hoje = date.today()
    prox_evento = None
    for dt, evt in agenda:
        if dt >= hoje:
            prox_evento = (dt, evt)
            break
    
    if prox_evento:
        p_dt, p_evt = prox_evento
        dias_falta = (p_dt - hoje).days
        txt_dias = "HOJE!" if dias_falta == 0 else f"Faltam {dias_falta} dias" if dias_falta > 0 else ""
        st.markdown(f"""
        <div class="next-event-box">
            <div class="next-label">✨ Próximo Ensaio • {txt_dias}</div>
            <div class="next-title">{p_evt['titulo']}</div>
            <div>{p_evt['local']}</div>
            <div class="next-date">{p_dt.day}/{p_dt.month} • {DIAS_SEMANA_PT[int(p_dt.strftime('%w'))]} • {p_evt['hora']}</div>
        </div>
        """, unsafe_allow_html=True)

    if not agenda:
        st.info("Nenhum evento encontrado.")
    else:
        mes_atual = 0
        for dt, evt_data in agenda:
            if dt.month != mes_atual:
                mes_atual = dt.month
                st.markdown(f"<div class='month-separator'><span class='month-text'>{NOMES_MESES[mes_atual]} {dt.year}</span></div>", unsafe_allow_html=True)
                if mes_atual in st.session_state['avisos'] and st.session_state['avisos'][mes_atual]:
                    aviso = st.session_state['avisos'][mes_atual]
                    st.markdown(f"""<div class="aviso-card"><span style="font-size: 20px">📢</span><span>{aviso}</span></div>""", unsafe_allow_html=True)
            
            dia_semana = DIAS_SEMANA_PT[int(dt.strftime("%w"))]
            mes_abrev = NOMES_MESES[dt.month][:3]
            link_google = gerar_link_google(dt, evt_data)
            
            st.markdown(f"""
            <div class="event-card">
                <div class="date-box">
                    <span class="date-day">{dt.day}</span>
                    <span class="date-month">{mes_abrev}</span>
                </div>
                <div class="event-details">
                    <div class="event-badge">{dia_semana}</div>
                    <div class="event-title">{evt_data['titulo']}</div>
                    <div class="event-info">📍 {evt_data['local']}</div>
                    <div class="event-info">🕒 {evt_data['hora']}</div>
                    <a href="{link_google}" target="_blank" class="btn-notify">🔔 Lembrete</a>
                </div>
            </div>
            """, unsafe_allow_html=True)

elif st.session_state['nav'] == 'Admin':
    st.markdown("<div class='admin-container'>", unsafe_allow_html=True)
    st.markdown("<h2 class='admin-title'>🔒 Painel Administrativo</h2>", unsafe_allow_html=True)
    
    senha = st.text_input("Senha de Acesso", type="password")
    if senha == "ccb123":
        st.success("✅ Acesso Liberado")
        st.session_state['ano_base'] = st.number_input("Ano de Referência", value=st.session_state['ano_base'], step=1)
        st.markdown("---")
        abas = st.tabs(["➕ Novo Evento", "📝 Avisos", "📋 Gerenciar", "📥 Downloads"])
        with abas[0]: 
            with st.form("add"):
                nome = st.text_input("Nome", "ENSAIO LOCAL")
                local = st.text_input("Local")
                dia = st.selectbox("Dia", [0,1,2,3,4,5,6], format_func=lambda x: DIAS_SEMANA_PT[x])
                semana = st.selectbox("Semana", ["1","2","3","4","5"])
                hora = st.text_input("Hora", "19:30 HRS")
                interc = st.selectbox("Frequência", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
                if st.form_submit_button("Salvar Evento"):
                    st.session_state['eventos'].append({"nome": nome.upper(), "local": local.upper(), "dia_sem": str(dia), "semana": semana, "hora": hora.upper(), "interc": interc})
                    st.rerun()
        with abas[1]: 
            mes_aviso = st.selectbox("Escolha o Mês", range(1, 13), format_func=lambda x: NOMES_MESES[x])
            texto_atual = st.session_state['avisos'].get(mes_aviso, "")
            novo_aviso = st.text_area("Texto do Aviso", value=texto_atual, height=100)
            c1, c2 = st.columns(2)
            if c1.button("Salvar Aviso"):
                st.session_state['avisos'][mes_aviso] = novo_aviso
                st.rerun()
            if c2.button("Apagar Aviso"):
                if mes_aviso in st.session_state['avisos']: del st.session_state['avisos'][mes_aviso]
                st.rerun()
        with abas[2]: 
            for i, evt in enumerate(st.session_state['eventos']):
                c_a, c_b = st.columns([4,1])
                c_a.write(f"**{evt['local']}** - {evt['semana']}ª {DIAS_SEMANA_CURTO[int(evt['dia_sem'])]}")
                if c_b.button("🗑️", key=f"d{i}"):
                    st.session_state['eventos'].pop(i)
                    st.rerun()
        with abas[3]: 
            d_excel = gerar_excel_todos_meses(st.session_state['ano_base'], st.session_state['eventos'], st.session_state['avisos'])
            st.download_button("⬇️ Excel", d_excel, f"Calendario_{st.session_state['ano_base']}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            d_pdf = gerar_pdf_calendario(st.session_state['ano_base'], st.session_state['eventos'], st.session_state['avisos'])
            st.download_button("⬇️ PDF", d_pdf, f"Calendario_{st.session_state['ano_base']}.pdf", mime="application/pdf")
    elif senha: st.error("❌ Senha Incorreta")
    st.markdown("</div>", unsafe_allow_html=True)
