import streamlit as st
import xlsxwriter
import calendar
from io import BytesIO
from datetime import datetime, date
from fpdf import FPDF
from urllib.parse import quote

# ==========================================
# 1. LÓGICA E FUNÇÕES
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

# ===== FUNÇÃO EXCEL (COM AVISOS) =====
def gerar_excel_todos_meses(ano, lista_eventos, avisos):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Calendário")
    
    header_mes = wb.add_format({'bold': True, 'font_size': 14, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_dias = wb.add_format({'bold': True, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    cell_dia = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 'bold': True})
    
    cell_evento = wb.add_format({
        'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 8, 
        'text_wrap': True, 'bg_color': '#FFFF00', 'bold': True
    })
    
    cell_aviso = wb.add_format({
        'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 
        'text_wrap': True, 'bg_color': '#FFCDD2', 'bold': True, 'font_color': '#B71C1C'
    })
    
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
                if dia == 0:
                    ws.write(current_row, col, '', cell_vazio)
                else:
                    chave = f"{ano}-{mes}-{dia}"
                    if chave in eventos_dict:
                        texto = f"{dia}\n"
                        for evt in eventos_dict[chave]: texto += f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}\n"
                        ws.write(current_row, col, texto, cell_evento)
                    else:
                        ws.write(current_row, col, dia, cell_dia)
            current_row += 1
        
        # Campo Anotações com o Aviso do Mês
        aviso_mes = avisos.get(mes, "")
        texto_anotacao = f"Anotações: {aviso_mes}"
        
        current_row += 1
        ws.merge_range(current_row, 0, current_row, 6, texto_anotacao, 
                       cell_aviso if aviso_mes else wb.add_format({'border': 1, 'align': 'left'}))
        current_row += 2
    
    wb.close()
    output.seek(0)
    return output

# ===== FUNÇÃO PDF (COM AVISOS) =====
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
        for dia in DIAS_SEMANA_CURTO:
            pdf.cell(col_width, header_height, dia, 1, 0, 'C', fill=True)
        pdf.ln(header_height)
        
        cal_matrix = calendar.monthcalendar(ano, mes)
        y_start = pdf.get_y()
        
        for semana in cal_matrix:
            x_current = margin_left
            for dia in semana:
                chave = f"{ano}-{mes}-{dia}"
                
                if dia == 0:
                    pdf.set_fill_color(230, 230, 230)
                elif chave in eventos_dict:
                    pdf.set_fill_color(255, 255, 0)
                else:
                    pdf.set_fill_color(255, 255, 255)
                
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
                        for evt in eventos_dict[chave]:
                            texto += f"{evt['titulo']}\n{evt['local']}\n{evt['hora']}\n"
                        pdf.multi_cell(col_width - 2, 3, texto, 0, 'L')

                x_current += col_width
            y_start += row_height
            
        # Anotações / Avisos
        aviso_mes = avisos.get(mes, "")
        
        pdf.set_xy(margin_left, 260)
        pdf.set_font("Arial", "B", 10)
        pdf.set_text_color(0, 0, 0)
        
        # Se tiver aviso, pinta de vermelho claro pra destacar
        if aviso_mes:
            pdf.set_fill_color(255, 230, 230) # Vermelho claro
            pdf.cell(190, 6, "Anotacoes / Avisos Importantes:", "LTR", 1, 'L', fill=True)
            pdf.set_font("Arial", "B", 11)
            pdf.set_text_color(180, 0, 0) # Texto vermelho escuro
            pdf.multi_cell(190, 15, aviso_mes, "LBR", 'L', fill=True)
        else:
            pdf.set_fill_color(255, 255, 255)
            pdf.cell(190, 6, "Anotacoes:", "LTR", 1, 'L')
            pdf.cell(190, 15, "", "LBR", 1, 'L')

    try:
        val = pdf.output(dest='S')
        if isinstance(val, str): return val.encode('latin-1')
        return bytes(val)
    except:
        return bytes(pdf.output())

# ==========================================
# 2. VISUAL DO APP
# ==========================================
st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    .stApp { background-color: #F5F7FA; }
    .block-container { padding-top: 0rem; padding-bottom: 6rem; padding-left: 1rem; padding-right: 1rem; max-width: 100%; }
    .app-header { position: fixed; top: 0; left: 0; width: 100%; background-color: #1F4E5F; color: white; padding: 15px 20px; z-index: 999; box-shadow: 0 2px 5px rgba(0,0,0,0.2); display: flex; align-items: center; justify-content: center; }
    .app-header h1 { margin: 0; font-family: 'Roboto', sans-serif; font-size: 18px; font-weight: 600; color: white !important; }
    .header-spacer { height: 70px; }
    .event-card { background: white; border-radius: 16px; padding: 16px; margin-bottom: 14px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid rgba(0,0,0,0.03); display: flex; align-items: flex-start; gap: 15px; }
    .date-box { background-color: #EBF2F5; border-radius: 12px; min-width: 60px; height: 60px; display: flex; flex-direction: column; align-items: center; justify-content: center; color: #1F4E5F; }
    .date-day { font-size: 22px; font-weight: 800; line-height: 1; }
    .date-month { font-size: 10px; font-weight: 600; text-transform: uppercase; margin-top: 2px; }
    .event-details { flex-grow: 1; }
    .event-badge { background-color: #1F4E5F; color: white; font-size: 9px; font-weight: 700; padding: 2px 8px; border-radius: 10px; text-transform: uppercase; display: inline-block; margin-bottom: 4px; }
    .event-title { font-size: 15px; font-weight: 700; color: #222; margin: 4px 0; line-height: 1.3; }
    .event-info { font-size: 13px; color: #666; display: flex; align-items: center; gap: 6px; margin-top: 2px; }
    .btn-notify { display: inline-block; margin-top: 8px; background-color: #eef6f8; color: #1F4E5F; font-size: 11px; font-weight: bold; padding: 6px 12px; border-radius: 20px; text-decoration: none; border: 1px solid #dbebf0; }
    .btn-notify:hover { background-color: #1F4E5F; color: white; border-color: #1F4E5F; }
    
    .month-separator { margin: 35px 0 20px 0; padding-left: 5px; }
    .month-text { 
        font-size: 24px !important; 
        font-weight: 900 !important; 
        color: #1F4E5F !important; 
        text-transform: uppercase; 
        letter-spacing: 1.5px; 
        border-bottom: 3px solid #A0C1D1; 
        padding-bottom: 5px;
        display: inline-block;
    }
    
    /* Estilo para o Aviso no App */
    .aviso-card {
        background-color: #FFEBEE;
        border-left: 5px solid #D32F2F;
        padding: 15px;
        margin-bottom: 15px;
        border-radius: 8px;
        color: #B71C1C;
        font-weight: bold;
        font-size: 14px;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .admin-container { background: white; padding: 20px; border-radius: 12px; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="app-header"><h1>AGENDA CCB JACIARA</h1></div><div class="header-spacer"></div>', unsafe_allow_html=True)

# --- DADOS INICIAIS ---
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

# --- NOVO ESTADO PARA AVISOS ---
if 'avisos' not in st.session_state:
    st.session_state['avisos'] = {} # Ex: {1: "Aviso Janeiro", 5: "Aviso Maio"}

if 'nav' not in st.session_state: st.session_state['nav'] = 'Agenda'
if 'ano_base' not in st.session_state: st.session_state['ano_base'] = date.today().year + 1

c1, c2 = st.columns(2)
with c1:
    if st.button("📅 Agenda", use_container_width=True): st.session_state['nav'] = 'Agenda'
with c2:
    if st.button("⚙️ Admin", use_container_width=True): st.session_state['nav'] = 'Admin'

if st.session_state['nav'] == 'Agenda':
    agenda = montar_agenda_ordenada(st.session_state['ano_base'], st.session_state['eventos'])
    if not agenda:
        st.info("Nenhum evento encontrado.")
    else:
        mes_atual = 0
        for dt, evt_data in agenda:
            if dt.month != mes_atual:
                mes_atual = dt.month
                st.markdown(f"<div class='month-separator'><span class='month-text'>{NOMES_MESES[mes_atual]} {dt.year}</span></div>", unsafe_allow_html=True)
                
                # --- EXIBIR AVISO SE HOUVER ---
                if mes_atual in st.session_state['avisos'] and st.session_state['avisos'][mes_atual]:
                    aviso = st.session_state['avisos'][mes_atual]
                    st.markdown(f"""
                    <div class="aviso-card">
                        <span>⚠️</span>
                        <span>{aviso}</span>
                    </div>
                    """, unsafe_allow_html=True)
            
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
                    <a href="{link_google}" target="_blank" class="btn-notify">🔔 Criar Lembrete</a>
                </div>
            </div>
            """, unsafe_allow_html=True)

elif st.session_state['nav'] == 'Admin':
    st.markdown("<div class='admin-container'>", unsafe_allow_html=True)
    st.subheader("Painel Administrativo")
    senha = st.text_input("Senha de Acesso", type="password")
    
    if senha == "ccb123":
        st.success("✅ Acesso Liberado")
        
        st.markdown("#### 🔧 Configurações Gerais")
        st.session_state['ano_base'] = st.number_input("Ano de Referência", value=st.session_state['ano_base'], step=1)
        uploaded_logo = st.file_uploader("Logo da Igreja (Para o Excel)", type=['jpg', 'png'])
        
        st.markdown("---")
        
        # --- NOVA ABA DE AVISOS ---
        abas = st.tabs(["➕ Novo Evento", "📝 Avisos/Observações", "📋 Gerenciar Eventos", "📥 Baixar Arquivos"])
        
        with abas[0]: # Novo Evento
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
        
        with abas[1]: # Avisos
            st.markdown("Adicione observações importantes para aparecerem no mês (ex: Mudança de data).")
            mes_aviso = st.selectbox("Escolha o Mês", range(1, 13), format_func=lambda x: NOMES_MESES[x])
            
            # Carrega aviso existente
            texto_atual = st.session_state['avisos'].get(mes_aviso, "")
            novo_aviso = st.text_area("Texto do Aviso", value=texto_atual, height=100)
            
            c_salvar, c_limpar = st.columns(2)
            if c_salvar.button("Salvar Aviso"):
                st.session_state['avisos'][mes_aviso] = novo_aviso
                st.success(f"Aviso de {NOMES_MESES[mes_aviso]} salvo!")
                st.rerun()
            
            if c_limpar.button("🗑️ Apagar Aviso"):
                if mes_aviso in st.session_state['avisos']:
                    del st.session_state['avisos'][mes_aviso]
                    st.success("Aviso removido.")
                    st.rerun()

        with abas[2]: # Gerenciar
            for i, evt in enumerate(st.session_state['eventos']):
                c_a, c_b = st.columns([4,1])
                c_a.write(f"**{evt['local']}** - {evt['semana']}ª {DIAS_SEMANA_CURTO[int(evt['dia_sem'])]}")
                if c_b.button("🗑️", key=f"d{i}"):
                    st.session_state['eventos'].pop(i)
                    st.rerun()
        
        with abas[3]: # Baixar
            st.write("**Calendário Excel (.xlsx)**")
            d_excel = gerar_excel_todos_meses(st.session_state['ano_base'], st.session_state['eventos'], st.session_state['avisos'])
            st.download_button(
                label="⬇️ Baixar Excel",
                data=d_excel,
                file_name=f"Calendario_{st.session_state['ano_base']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="excel_btn"
            )
            
            st.write("**Calendário PDF (.pdf)**")
            d_pdf = gerar_pdf_calendario(st.session_state['ano_base'], st.session_state['eventos'], st.session_state['avisos'])
            st.download_button(
                label="⬇️ Baixar PDF",
                data=d_pdf,
                file_name=f"Calendario_{st.session_state['ano_base']}.pdf",
                mime="application/pdf",
                key="pdf_btn"
            )

    elif senha:
        st.error("❌ Senha Incorreta")
    st.markdown("</div>", unsafe_allow_html=True)
