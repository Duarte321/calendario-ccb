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
    # Link otimizado para abrir no app do celular
    titulo = quote(f"{evt_data['titulo']} - {evt_data['local']}")
    hora_limpa = evt_data['hora'].replace("HRS", "").replace(":", "").strip()
    if len(hora_limpa) < 4: hora_limpa = "1930"
    data_inicio = f"{dt.year}{dt.month:02d}{dt.day:02d}T{hora_limpa}00"
    data_fim = f"{dt.year}{dt.month:02d}{dt.day:02d}T{int(hora_limpa[:2])+2:02d}{hora_limpa[2:]}00"
    local = quote(evt_data['local'])
    return f"https://calendar.google.com/calendar/r/eventedit?text={titulo}&dates={data_inicio}/{data_fim}&location={local}&details=Ensaio+CCB"

# ===== FUNÇÃO EXCEL COMPLETA (RESTAURADA) =====
def gerar_excel_buffer(ano, lista_eventos, uploaded_logo=None):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Estilos Profissionais
    header_fmt = wb.add_format({'bold': True, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'bg_color': '#EBF2F5', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    data_fmt = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    center_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    
    # Calcula eventos
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    
    # Cria uma aba para cada mês
    for mes in range(1, 13):
        nome_mes = NOMES_MESES[mes]
        ws = wb.add_worksheet(nome_mes)
        
        # Configuração de largura das colunas
        ws.set_column('A:A', 12) # Data
        ws.set_column('B:B', 10) # Dia Sem
        ws.set_column('C:C', 30) # Evento
        ws.set_column('D:D', 35) # Local
        ws.set_column('E:E', 12) # Hora
        
        # Cabeçalho com Título e Logo
        if uploaded_logo:
            try:
                img_data = BytesIO(uploaded_logo)
                ws.insert_image('A1', 'logo.png', {'image_data': img_data, 'x_scale': 0.08, 'y_scale': 0.08})
                ws.merge_range('A1:E2', f"AGENDA - {nome_mes} {ano}", title_fmt)
            except:
                ws.merge_range('A1:E1', f"{nome_mes} {ano}", title_fmt)
        else:
            ws.merge_range('A1:E1', f"{nome_mes} {ano}", title_fmt)
        
        # Linha de Títulos das Colunas
        row_header = 2 if uploaded_logo else 1
        ws.write(row_header, 0, 'DATA', header_fmt)
        ws.write(row_header, 1, 'DIA', header_fmt)
        ws.write(row_header, 2, 'EVENTO', header_fmt)
        ws.write(row_header, 3, 'LOCAL', header_fmt)
        ws.write(row_header, 4, 'HORA', header_fmt)
        
        # Preenchimento dos Dados
        linha = row_header + 1
        for dt, evt_data in agenda:
            if dt.month == mes and dt.year == ano:
                ws.write(linha, 0, dt.strftime('%d/%m/%Y'), center_fmt)
                ws.write(linha, 1, DIAS_SEMANA_CURTO[int(dt.strftime("%w"))], center_fmt)
                ws.write(linha, 2, evt_data['titulo'], data_fmt)
                ws.write(linha, 3, evt_data['local'], data_fmt)
                ws.write(linha, 4, evt_data['hora'], center_fmt)
                linha += 1
    
    wb.close()
    output.seek(0)
    return output

# ===== FUNÇÃO PDF COMPLETA (RESTAURADA E CORRIGIDA) =====
def gerar_pdf_buffer(ano, lista_eventos):
    # Configuração PDF
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Título Principal
    pdf.set_font("Arial", "B", 20)
    pdf.set_text_color(31, 78, 95) # Cor Azul Petróleo
    pdf.cell(0, 15, f"CALENDARIO CCB JACIARA {ano}", 0, 1, 'C')
    
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)
    
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    
    mes_atual = 0
    for dt, evt_data in agenda:
        if dt.month != mes_atual:
            mes_atual = dt.month
            pdf.ln(5)
            
            # Cabeçalho do Mês com fundo colorido (simulando o Excel)
            pdf.set_fill_color(235, 242, 245) # Cinza azulado claro
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(31, 78, 95)
            
            # Tratamento de acentuação do Mês
            nome_mes = NOMES_MESES[mes_atual]
            try:
                nome_mes = nome_mes.encode('latin-1', 'ignore').decode('latin-1')
            except:
                pass
            
            pdf.cell(0, 8, f"  {nome_mes} {ano}", 0, 1, 'L', fill=True)
            pdf.line(10, pdf.get_y(), 200, pdf.get_y()) # Linha abaixo do mês
            pdf.ln(2)
            pdf.set_font("Arial", "", 10)
            pdf.set_text_color(0, 0, 0)
        
        # Tratamento de texto dos eventos
        try:
            dia_semana = DIAS_SEMANA_PT[int(dt.strftime("%w"))].encode('latin-1', 'ignore').decode('latin-1')
            titulo = evt_data['titulo'].encode('latin-1', 'ignore').decode('latin-1')
            local = evt_data['local'].encode('latin-1', 'ignore').decode('latin-1')
        except:
            dia_semana = DIAS_SEMANA_PT[int(dt.strftime("%w"))]
            titulo = evt_data['titulo']
            local = evt_data['local']

        # Linha do evento
        texto = f"{dt.day:02d}/{dt.month:02d} ({dia_semana}) - {titulo} - {local} ({evt_data['hora']})"
        pdf.multi_cell(0, 6, texto, 0, 'L')
        
        # Linha divisória sutil entre eventos
        x_start = pdf.get_x()
        y_start = pdf.get_y()
        pdf.set_draw_color(220, 220, 220)
        pdf.line(x_start, y_start, 200, y_start)
        pdf.set_draw_color(0, 0, 0) # Reseta cor preta
    
    # Solução para o erro AttributeError
    val = pdf.output(dest='S')
    if isinstance(val, str):
        return val.encode('latin-1')
    return val

# ==========================================
# 2. VISUAL DO APP (INTERFACE)
# ==========================================
st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp { background-color: #F5F7FA; }
    .block-container { padding-top: 0rem; padding-bottom: 6rem; padding-left: 1rem; padding-right: 1rem; max-width: 100%; }

    /* HEADER */
    .app-header {
        position: fixed; top: 0; left: 0; width: 100%;
        background-color: #1F4E5F; color: white;
        padding: 15px 20px; z-index: 999;
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
        display: flex; align-items: center; justify-content: center;
    }
    .app-header h1 { margin: 0; font-family: 'Roboto', sans-serif; font-size: 18px; font-weight: 600; color: white !important; }
    .header-spacer { height: 70px; }

    /* CARTÕES */
    .event-card {
        background: white; border-radius: 16px; padding: 16px; margin-bottom: 14px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid rgba(0,0,0,0.03);
        display: flex; align-items: flex-start; gap: 15px; position: relative;
    }
    .date-box {
        background-color: #EBF2F5; border-radius: 12px; min-width: 60px; height: 60px;
        display: flex; flex-direction: column; align-items: center; justify-content: center; color: #1F4E5F;
    }
    .date-day { font-size: 22px; font-weight: 800; line-height: 1; }
    .date-month { font-size: 10px; font-weight: 600; text-transform: uppercase; margin-top: 2px; }
    .event-details { flex-grow: 1; }
    .event-badge {
        background-color: #1F4E5F; color: white; font-size: 9px; font-weight: 700;
        padding: 2px 8px; border-radius: 10px; text-transform: uppercase; display: inline-block; margin-bottom: 4px;
    }
    .event-title { font-size: 15px; font-weight: 700; color: #222; margin: 4px 0; line-height: 1.3; }
    .event-info { font-size: 13px; color: #666; display: flex; align-items: center; gap: 6px; margin-top: 2px; }
    
    /* BOTÃO DE NOTIFICAÇÃO */
    .btn-notify {
        display: inline-block;
        margin-top: 8px;
        background-color: #eef6f8;
        color: #1F4E5F;
        font-size: 11px;
        font-weight: bold;
        padding: 6px 12px;
        border-radius: 20px;
        text-decoration: none;
        border: 1px solid #dbebf0;
    }
    .btn-notify:hover { background-color: #1F4E5F; color: white; border-color: #1F4E5F; }

    /* DIVISOR MÊS */
    .month-separator { margin: 25px 0 15px 0; padding-left: 5px; }
    .month-text { font-size: 14px; font-weight: 800; color: #8898AA; text-transform: uppercase; letter-spacing: 1px; }
    
    /* ADMIN */
    .admin-container { background: white; padding: 20px; border-radius: 12px; margin-top: 10px; }
</style>
""", unsafe_allow_html=True)

# --- LAYOUT ---
st.markdown('<div class="app-header"><h1>AGENDA CCB JACIARA</h1></div><div class="header-spacer"></div>', unsafe_allow_html=True)

# Dados Iniciais
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

if 'nav' not in st.session_state: st.session_state['nav'] = 'Agenda'
if 'ano_base' not in st.session_state: st.session_state['ano_base'] = date.today().year + 1

# Botões de Navegação
c1, c2 = st.columns(2)
with c1:
    if st.button("📅 Agenda", use_container_width=True): st.session_state['nav'] = 'Agenda'
with c2:
    if st.button("⚙️ Admin", use_container_width=True): st.session_state['nav'] = 'Admin'

# --- PÁGINA AGENDA ---
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

# --- PÁGINA ADMIN ---
elif st.session_state['nav'] == 'Admin':
    st.markdown("<div class='admin-container'>", unsafe_allow_html=True)
    st.subheader("Painel Administrativo")
    senha = st.text_input("Senha de Acesso", type="password")
    
    if senha == "ccb123":
        st.success("✅ Acesso Liberado")
        
        # 1. Configurações Gerais (Ano e Logo)
        st.markdown("#### 🔧 Configurações Gerais")
        st.session_state['ano_base'] = st.number_input("Ano de Referência", value=st.session_state['ano_base'], step=1)
        uploaded_logo = st.file_uploader("Logo da Igreja (Para o Excel)", type=['jpg', 'png'])
        logo_data = uploaded_logo.getvalue() if uploaded_logo else None
        
        st.markdown("---")
        
        # 2. Novo Evento
        with st.expander("➕ Cadastrar Novo Evento"):
            with st.form("add"):
                nome = st.text_input("Nome", "ENSAIO LOCAL")
                local = st.text_input("Local")
                dia = st.selectbox("Dia", [0,1,2,3,4,5,6], format_func=lambda x: DIAS_SEMANA_PT[x])
                semana = st.selectbox("Semana", ["1","2","3","4","5"])
                hora = st.text_input("Hora", "19:30 HRS")
                interc = st.selectbox("Frequência", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
                if st.form_submit_button("Salvar"):
                    st.session_state['eventos'].append({"nome": nome.upper(), "local": local.upper(), "dia_sem": str(dia), "semana": semana, "hora": hora.upper(), "interc": interc})
                    st.rerun()
        
        st.markdown("---")
        
        # 3. Gerenciamento de Eventos
        st.markdown("#### 📋 Gerenciar Eventos")
        for i, evt in enumerate(st.session_state['eventos']):
            c_a, c_b = st.columns([4,1])
            c_a.write(f"**{evt['local']}** - {evt['semana']}ª {DIAS_SEMANA_CURTO[int(evt['dia_sem'])]}")
            if c_b.button("🗑️", key=f"d{i}"):
                st.session_state['eventos'].pop(i)
                st.rerun()
        
        st.markdown("---")
        
        # 4. Exportação (PDF e Excel)
        st.markdown("#### 📥 Baixar Arquivos")
        c_exc, c_pdf = st.columns(2)
        
        with c_exc:
            st.write("**Excel (.xlsx)**")
            d_excel = gerar_excel_buffer(st.session_state['ano_base'], st.session_state['eventos'], logo_data)
            st.download_button("⬇️ Baixar Excel", d_excel, f"Calendario_{st.session_state['ano_base']}.xlsx", key="excel_btn")
        
        with c_pdf:
            st.write("**PDF (.pdf)**")
            d_pdf = gerar_pdf_buffer(st.session_state['ano_base'], st.session_state['eventos'])
            st.download_button("⬇️ Baixar PDF", d_pdf, f"Calendario_{st.session_state['ano_base']}.pdf", key="pdf_btn")

    elif senha:
        st.error("❌ Senha Incorreta")
    
    st.markdown("</div>", unsafe_allow_html=True)
