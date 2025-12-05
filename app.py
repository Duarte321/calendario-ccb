import streamlit as st
import xlsxwriter
import calendar
from io import BytesIO
from datetime import datetime, date
from fpdf import FPDF

# ==========================================
# 1. LÓGICA DO CALENDÁRIO
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
                    evento_dados = {
                        "titulo": evt['nome'],
                        "local": evt['local'],
                        "hora": evt['hora']
                    }
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

def gerar_excel_buffer(ano, lista_eventos, uploaded_logo):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet(f"Calendário {ano}")

    # Estilos Excel
    COR_VERDE_ESCURO = '#1F4E5F'
    COR_AMARELO_NEON = '#FFFF00'
    COR_CINZA_LINHA  = '#D9D9D9'
    fmt_ano = wb.add_format({'bold': True, 'font_size': 24, 'font_color': 'white', 'bg_color': COR_VERDE_ESCURO, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_mes_nome = wb.add_format({'font_size': 28, 'font_color': COR_VERDE_ESCURO, 'align': 'left', 'valign': 'bottom'})
    fmt_header_sem = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': COR_VERDE_ESCURO, 'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'border': 0})
    fmt_dia_box = wb.add_format({'valign': 'top', 'align': 'left', 'border': 1, 'border_color': COR_CINZA_LINHA, 'font_size': 11})
    fmt_evento_bg = wb.add_format({'valign': 'center', 'align': 'center', 'border': 1, 'border_color': COR_CINZA_LINHA, 'bg_color': COR_AMARELO_NEON, 'text_wrap': True, 'font_size': 10, 'bold': True})
    fmt_logo_celula = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

    dados_agenda = calcular_eventos(ano, lista_eventos)
    dados_simples = {}
    for k, v_list in dados_agenda.items():
        textos = [f"{x['titulo']} {x['local']} - {x['hora']}" for x in v_list]
        dados_simples[k] = textos

    calendar.setfirstweekday(calendar.SUNDAY)
    LINHA = 0
    for mes in range(1, 13):
        ws.write(LINHA, 0, ano, fmt_ano)
        ws.merge_range(LINHA, 1, LINHA, 5, NOMES_MESES[mes], fmt_mes_nome)
        ws.set_row(LINHA, 40)
        if uploaded_logo:
            ws.insert_image(LINHA, 6, "logo.jpg", {'image_data': uploaded_logo, 'x_scale': 0.25, 'y_scale': 0.25, 'x_offset': 5, 'y_offset': 2, 'positioning': 2})
        else:
            ws.write(LINHA, 6, "", fmt_logo_celula)
        LINHA += 1
        ws.write_row(LINHA, 0, DIAS_SEMANA_PT, fmt_header_sem)
        LINHA += 1
        cal = calendar.monthcalendar(ano, mes)
        for semana in cal:
            ws.set_row(LINHA, 60)
            COL = 0
            for dia in semana:
                if dia == 0:
                    ws.write(LINHA, COL, "", fmt_dia_box)
                else:
                    chave = f"{ano}-{mes}-{dia}"
                    if chave in dados_simples:
                        textos_evt = "\n".join(dados_simples[chave])
                        ws.write(LINHA, COL, f"{dia}\n{textos_evt}", fmt_evento_bg)
                    else:
                        ws.write(LINHA, COL, dia, fmt_dia_box)
                COL += 1
            LINHA += 1
        ws.merge_range(LINHA, 0, LINHA, 6, " Anotações:", fmt_dia_box)
        LINHA += 2
    ws.set_column('A:G', 18)
    wb.close()
    output.seek(0)
    return output

def gerar_pdf_buffer(ano, lista_eventos):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False, margin=8)
    
    dados_agenda = calcular_eventos(ano, lista_eventos)
    dados_simples = {}
    for k, v_list in dados_agenda.items():
        textos = [f"{x['titulo']}\n{x['local']} AS {x['hora']}" for x in v_list]
        dados_simples[k] = textos
    
    for mes in range(1, 13):
        pdf.add_page()
        pdf.set_font("Helvetica", style="B", size=20)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(31, 78, 95)
        pdf.cell(25, 12, str(ano), border=1, align="C", fill=True)
        mes_nome = NOMES_MESES[mes].lower()
        pdf.cell(0, 12, mes_nome, border=1, align="C", fill=True, ln=1)
        pdf.ln(3)
        pdf.set_font("Helvetica", style="B", size=8)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(31, 78, 95)
        dias_semana_pt = ["DOMINGO", "SEGUNDA-FEIRA", "TERCA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SABADO"]
        largura_coluna = (pdf.w - 16) / 7
        altura_header = 8
        for dia_sem in dias_semana_pt:
            pdf.cell(largura_coluna, altura_header, dia_sem, border=1, align="C", fill=True)
        pdf.ln()
        calendar.setfirstweekday(calendar.SUNDAY)
        cal = calendar.monthcalendar(ano, mes)
        altura_dia = 32
        for semana in cal:
            y_inicial = pdf.get_y()
            x_inicial = pdf.get_x()
            for idx, dia in enumerate(semana):
                pdf.set_xy(x_inicial + (idx * largura_coluna), y_inicial)
                if dia == 0:
                    pdf.set_fill_color(230, 230, 230)
                    pdf.cell(largura_coluna, altura_dia, "", border=1, fill=True)
                else:
                    chave = f"{ano}-{mes}-{dia}"
                    if chave in dados_simples:
                        pdf.set_fill_color(255, 255, 0)
                        pdf.cell(largura_coluna, altura_dia, "", border=1, fill=True)
                    else:
                        pdf.set_fill_color(255, 255, 255)
                        pdf.cell(largura_coluna, altura_dia, "", border=1, fill=True)
            for idx, dia in enumerate(semana):
                if dia != 0:
                    pdf.set_xy(x_inicial + (idx * largura_coluna), y_inicial)
                    pdf.set_font("Helvetica", style="B", size=11)
                    pdf.set_text_color(0, 0, 0)
                    pdf.set_xy(x_inicial + (idx * largura_coluna) + 1, y_inicial + 1)
                    pdf.cell(6, 5, str(dia), border=0, align="L")
                    chave = f"{ano}-{mes}-{dia}"
                    if chave in dados_simples:
                        texto_evento = "\n".join(dados_simples[chave])
                        texto_limpo = texto_evento.replace("Á", "A").replace("á", "a").replace("É", "E").replace("é", "e").replace("Í", "I").replace("í", "i").replace("Ó", "O").replace("ó", "o").replace("Ú", "U").replace("ú", "u").replace("Ã", "A").replace("ã", "a").replace("Õ", "O").replace("õ", "o").replace("Ç", "C").replace("ç", "c")
                        pdf.set_xy(x_inicial + (idx * largura_coluna), y_inicial + 7)
                        pdf.set_font("Helvetica", style="B", size=8)
                        pdf.multi_cell(largura_coluna, 3.5, texto_limpo, align="C")
            pdf.set_xy(x_inicial, y_inicial + altura_dia)
        pdf.ln(3)
        pdf.set_font("Helvetica", style="", size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(0, 6, " Anotacoes:", border=1, ln=1)
    return bytes(pdf.output())

# ==========================================
# 2. INTERFACE (MOBILE APP LOOK)
# ==========================================
st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered", initial_sidebar_state="collapsed")

# CSS AVANÇADO PARA VISUAL DE APP NATIVO
st.markdown("""
<style>
    /* Ocultar elementos padrão do Streamlit para parecer App */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Fundo geral */
    .stApp {
        background-color: #F5F7FA;
    }
    
    /* Área de conteúdo principal */
    .block-container {
        padding-top: 0rem;
        padding-bottom: 6rem; /* Espaço para a barra inferior */
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 100%;
    }

    /* HEADER FIXO NO TOPO */
    .app-header {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        background-color: #1F4E5F;
        color: white;
        padding: 15px 20px;
        z-index: 999;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .app-header h1 {
        margin: 0;
        font-family: 'Roboto', sans-serif;
        font-size: 18px;
        font-weight: 600;
        color: white !important;
        letter-spacing: 0.5px;
    }
    
    /* Espaçador para compensar o header fixo */
    .header-spacer {
        height: 70px;
    }

    /* ESTILO DOS CARTÕES (Card View) */
    .event-card {
        background: white;
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        border: 1px solid rgba(0,0,0,0.03);
        display: flex;
        align-items: flex-start;
        gap: 15px;
    }
    
    /* Coluna da Data (Esquerda) */
    .date-box {
        background-color: #EBF2F5;
        border-radius: 12px;
        min-width: 60px;
        height: 60px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        color: #1F4E5F;
    }
    .date-day { font-size: 22px; font-weight: 800; line-height: 1; }
    .date-month { font-size: 10px; font-weight: 600; text-transform: uppercase; margin-top: 2px; }

    /* Coluna de Detalhes (Direita) */
    .event-details { flex-grow: 1; }
    .event-badge {
        background-color: #1F4E5F;
        color: white;
        font-size: 9px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 10px;
        text-transform: uppercase;
        display: inline-block;
        margin-bottom: 4px;
        letter-spacing: 0.5px;
    }
    .event-title {
        font-size: 15px;
        font-weight: 700;
        color: #222;
        margin: 4px 0;
        line-height: 1.3;
    }
    .event-info {
        font-size: 13px;
        color: #666;
        display: flex;
        align-items: center;
        gap: 6px;
        margin-top: 2px;
    }

    /* Divisor de Mês Elegante */
    .month-separator {
        margin: 25px 0 15px 0;
        padding-left: 5px;
    }
    .month-text {
        font-size: 14px;
        font-weight: 800;
        color: #8898AA;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    /* BARRA DE NAVEGAÇÃO INFERIOR FIXA */
    .bottom-nav {
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
        background-color: white;
        border-top: 1px solid #eee;
        display: flex;
        justify-content: space-around;
        padding: 10px 0;
        z-index: 999;
        box-shadow: 0 -2px 10px rgba(0,0,0,0.03);
    }
    .nav-item {
        text-align: center;
        cursor: pointer;
        width: 50%;
    }
    .nav-icon { font-size: 20px; display: block; margin-bottom: 2px; }
    .nav-label { font-size: 10px; font-weight: 600; color: #888; }
    
    /* Botão Selecionado na Nav */
    .nav-active .nav-icon, .nav-active .nav-label { color: #1F4E5F; }
    
    /* Ajustes Admin */
    .admin-container {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        margin-top: 10px;
    }

</style>
""", unsafe_allow_html=True)

# --- HEADER FIXO ---
st.markdown("""
<div class="app-header">
    <h1>AGENDA CCB JACIARA</h1>
</div>
<div class="header-spacer"></div>
""", unsafe_allow_html=True)

# --- ESTADO E DADOS ---
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

# Gerenciamento de Navegação (Estado)
if 'nav' not in st.session_state:
    st.session_state['nav'] = 'Agenda'

# Botões invisíveis para controlar navegação (Hack para simular Tab Bar)
c1, c2 = st.columns(2)
with c1:
    if st.button("📅 Agenda", use_container_width=True): st.session_state['nav'] = 'Agenda'
with c2:
    if st.button("⚙️ Admin", use_container_width=True): st.session_state['nav'] = 'Admin'

# --- PÁGINA: AGENDA (HOME) ---
if st.session_state['nav'] == 'Agenda':
    
    ano_atual = date.today().year + 1
    agenda = montar_agenda_ordenada(ano_atual, st.session_state['eventos'])
    
    if not agenda:
        st.info("Nenhum evento encontrado.")
    else:
        mes_atual = 0
        for dt, evt_data in agenda:
            # Divisor de Mês
            if dt.month != mes_atual:
                mes_atual = dt.month
                nome_mes = NOMES_MESES[mes_atual]
                st.markdown(f"""
                <div class="month-separator">
                    <span class="month-text">{nome_mes} {dt.year}</span>
                </div>
                """, unsafe_allow_html=True)
            
            dia_semana = DIAS_SEMANA_PT[int(dt.strftime("%w"))]
            dia_num = dt.day
            mes_abrev = NOMES_MESES[dt.month][:3]
            
            # Cartão
            st.markdown(f"""
            <div class="event-card">
                <div class="date-box">
                    <span class="date-day">{dia_num}</span>
                    <span class="date-month">{mes_abrev}</span>
                </div>
                <div class="event-details">
                    <div class="event-badge">{dia_semana}</div>
                    <div class="event-title">{evt_data['titulo']}</div>
                    <div class="event-info">📍 {evt_data['local']}</div>
                    <div class="event-info">🕒 {evt_data['hora']}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
    st.write("") # Espaço extra no final

# --- PÁGINA: ADMIN ---
elif st.session_state['nav'] == 'Admin':
    
    st.markdown("<div class='admin-container'>", unsafe_allow_html=True)
    st.subheader("Acesso Restrito")
    
    senha = st.text_input("Senha", type="password", placeholder="Digite a senha")
    
    if senha == "ccb123":
        st.success("Logado com sucesso")
        
        st.markdown("#### Configurações")
        ano_escolhido = st.number_input("Ano de Referência", value=date.today().year + 1)
        uploaded_file = st.file_uploader("Logo da Igreja", type=['jpg', 'png'])
        logo_data = uploaded_file.getvalue() if uploaded_file else None
        
        st.markdown("---")
        st.markdown("#### Novo Evento")
        
        with st.form("form_add"):
            novo_nome = st.text_input("Nome do Evento", value="ENSAIO LOCAL")
            c_loc, c_hr = st.columns(2)
            with c_loc: novo_local = st.text_input("Local", placeholder="Ex: Jaciara")
            with c_hr: novo_hora = st.text_input("Hora", value="19:30 HRS")
            
            c_sem, c_dia = st.columns(2)
            with c_sem: novo_semana = st.selectbox("Semana", ["1", "2", "3", "4", "5"])
            with c_dia: novo_dia = st.selectbox("Dia", options=[0,1,2,3,4,5,6], format_func=lambda x: DIAS_SEMANA_PT[x].title(), index=5)
            
            novo_interc = st.selectbox("Repetição", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
            
            if st.form_submit_button("Salvar Evento", type="primary"):
                item = {"nome": novo_nome.upper(), "local": novo_local.upper(), "dia_sem": str(novo_dia), "semana": novo_semana, "hora": novo_hora.upper(), "interc": novo_interc}
                st.session_state['eventos'].append(item)
                st.success("Adicionado!")
                st.rerun()
        
        st.markdown("---")
        st.markdown("#### Eventos Ativos")
        for i, evt in enumerate(st.session_state['eventos']):
            with st.expander(f"{evt['nome']} - {evt['local']}"):
                st.write(f"{evt['semana']}ª {DIAS_SEMANA_PT[int(evt['dia_sem'])]}")
                if st.button("Excluir", key=f"del_{i}"):
                    st.session_state['eventos'].pop(i)
                    st.rerun()
        
        st.markdown("---")
        st.markdown("#### Exportar")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Gerar Excel"):
                d_excel = gerar_excel_buffer(ano_escolhido, st.session_state['eventos'], logo_data)
                st.download_button("Baixar .xlsx", d_excel, f"Calendario_{ano_escolhido}.xlsx")
        with col_b:
            if st.button("Gerar PDF"):
                d_pdf = gerar_pdf_buffer(ano_escolhido, st.session_state['eventos'])
                st.download_button("Baixar .pdf", d_pdf, f"Calendario_{ano_escolhido}.pdf")
                
    elif senha:
        st.error("Senha Incorreta")
    
    st.markdown("</div>", unsafe_allow_html=True)

# --- BARRA DE NAVEGAÇÃO VISUAL (CSS HACK) ---
# Como o Streamlit recarrega a página, usamos botões nativos no topo para lógica
# Mas visualmente injetamos uma barra fixa embaixo apenas para decoração/instrução se fosse SPA
# No Streamlit puro, a melhor navegação "app-like" é usar st.navigation ou os botões no topo como fiz acima.
