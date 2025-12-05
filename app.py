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
# 2. INTERFACE DO APP (STREAMLIT)
# ==========================================
st.set_page_config(page_title="Agenda CCB Jaciara", page_icon="📅", layout="centered")

# CSS APRIMORADO PARA VISUAL "PREMIUM CLEAN"
st.markdown("""
<style>
    /* Ajuste geral de fundo */
    .stApp {
        background-color: #f8f9fa;
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 4rem;
        max-width: 800px;
    }

    /* Estilo dos Cartões de Evento */
    .agenda-card {
        background-color: white;
        border-radius: 12px;
        padding: 0;
        margin-bottom: 16px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.06);
        border: 1px solid #e0e0e0;
        display: flex;
        overflow: hidden;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .agenda-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(0,0,0,0.1);
    }

    /* Coluna da Esquerda (Data) */
    .card-date-col {
        background-color: #F0F4F8; /* Cinza azulado bem claro */
        width: 85px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        border-right: 1px solid #e0e0e0;
        padding: 10px;
    }

    .date-number {
        font-size: 26px;
        font-weight: 800;
        color: #1F4E5F;
        line-height: 1;
    }

    .date-month {
        font-size: 10px;
        font-weight: 600;
        text-transform: uppercase;
        color: #666;
        margin-top: 4px;
    }

    /* Coluna da Direita (Conteúdo) */
    .card-content-col {
        padding: 16px;
        flex-grow: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }

    /* Badge do Dia da Semana */
    .weekday-badge {
        display: inline-block;
        background-color: #1F4E5F;
        color: white;
        font-size: 10px;
        font-weight: bold;
        padding: 3px 8px;
        border-radius: 4px;
        text-transform: uppercase;
        margin-bottom: 6px;
        width: fit-content;
    }

    .event-title {
        font-size: 16px;
        font-weight: 700;
        color: #222;
        margin-bottom: 6px;
        line-height: 1.3;
    }

    .event-meta {
        display: flex;
        align-items: center;
        gap: 12px;
        font-size: 13px;
        color: #555;
    }

    .event-meta-item {
        display: flex;
        align-items: center;
        gap: 4px;
    }

    /* Divisor de Mês */
    .month-divider {
        display: flex;
        align-items: center;
        margin: 30px 0 15px 0;
    }
    
    .month-label {
        background-color: #1F4E5F;
        color: white;
        padding: 6px 16px;
        border-radius: 20px;
        font-weight: bold;
        font-size: 14px;
        text-transform: uppercase;
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
    }
    
    .month-line {
        flex-grow: 1;
        height: 2px;
        background-color: #e0e0e0;
        margin-left: 12px;
    }

</style>
""", unsafe_allow_html=True)

# Título Principal
st.markdown("<h2 style='text-align: center; color: #1F4E5F; margin-bottom: 25px;'>📅 Agenda CCB Jaciara - MT</h2>", unsafe_allow_html=True)

# --- SIDEBAR ---
with st.sidebar:
    st.header("Painel de Controle")
    ano_escolhido = st.number_input("Ano", value=date.today().year + 1, step=1)
    uploaded_file = st.file_uploader("Logo (Opcional)", type=['jpg', 'png'])
    logo_data = uploaded_file.getvalue() if uploaded_file else None

# Inicialização de Dados
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

# --- SELETOR DE MODOS ---
col_mode_1, col_mode_2 = st.columns(2)
modo = st.radio("Menu de Acesso", ["Visualizar Agenda", "Área Administrativa"], horizontal=True, label_visibility="collapsed")

st.write("") # Espaçamento

if modo == "Área Administrativa":
    st.markdown("### 🔐 Acesso Restrito")
    senha = st.text_input("Senha de Acesso", type="password", placeholder="Digite a senha do encarregado")
    
    if senha == "ccb123":
        st.success("Acesso Autorizado")
        
        with st.expander("➕ Cadastrar Novo Evento", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                novo_nome = st.text_input("Descrição", value="ENSAIO LOCAL")
                novo_dia = st.selectbox("Dia", options=[0,1,2,3,4,5,6], format_func=lambda x: ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"][x], index=5)
                novo_interc = st.selectbox("Frequência", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
            with c2:
                novo_local = st.text_input("Localidade", placeholder="Ex: Central Jaciara")
                novo_semana = st.selectbox("Semana", options=["1", "2", "3", "4", "5"], index=0)
                novo_hora = st.text_input("Horário", value="19:30 HRS")
            
            if st.button("Salvar Evento", type="primary"):
                item = {"nome": novo_nome.upper(), "local": novo_local.upper(), "dia_sem": str(novo_dia), "semana": novo_semana, "hora": novo_hora.upper(), "interc": novo_interc}
                st.session_state['eventos'].append(item)
                st.rerun()

        st.markdown("---")
        st.subheader("Gerenciamento de Eventos")
        
        for i, evt in enumerate(st.session_state['eventos']):
            dia_txt = DIAS_SEMANA_CURTO[int(evt['dia_sem'])]
            with st.container():
                c_a, c_b = st.columns([5, 1])
                with c_a:
                    st.markdown(f"**{evt['nome']}** | {evt['local']}")
                    st.caption(f"{evt['semana']}ª {dia_txt} - {evt['hora']} ({evt['interc']})")
                with c_b:
                    if st.button("Excluir", key=f"del_{i}"):
                        st.session_state['eventos'].pop(i)
                        st.rerun()
            st.divider()

        st.markdown("### 📥 Exportar Dados")
        c_ex, c_pd = st.columns(2)
        with c_ex:
            if st.button("Baixar Excel"):
                excel_data = gerar_excel_buffer(ano_escolhido, st.session_state['eventos'], logo_data)
                st.download_button("Download .XLSX", excel_data, f"Agenda_{ano_escolhido}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c_pd:
            if st.button("Baixar PDF"):
                pdf_data = gerar_pdf_buffer(ano_escolhido, st.session_state['eventos'])
                st.download_button("Download .PDF", pdf_data, f"Agenda_{ano_escolhido}.pdf", "application/pdf")

    elif senha:
        st.error("Senha incorreta.")

else:
    # --- MODO AGENDA VISUAL ---
    agenda = montar_agenda_ordenada(ano_escolhido, st.session_state['eventos'])
    
    if not agenda:
        st.info("Nenhum evento agendado para este período.")
    else:
        mes_atual = 0
        for dt, evt_data in agenda:
            # Cabeçalho do Mês
            if dt.month != mes_atual:
                mes_atual = dt.month
                nome_mes = NOMES_MESES[mes_atual]
                st.markdown(f"""
                <div class="month-divider">
                    <div class="month-label">{nome_mes}</div>
                    <div class="month-line"></div>
                </div>
                """, unsafe_allow_html=True)
            
            dia_semana = DIAS_SEMANA_PT[int(dt.strftime("%w"))]
            dia_num = dt.day
            mes_abrev = NOMES_MESES[dt.month][:3]
            
            # Cartão do Evento HTML/CSS
            st.markdown(f"""
            <div class="agenda-card">
                <div class="card-date-col">
                    <div class="date-number">{dia_num}</div>
                    <div class="date-month">{mes_abrev}</div>
                </div>
                <div class="card-content-col">
                    <div class="weekday-badge">{dia_semana}</div>
                    <div class="event-title">{evt_data['titulo']}</div>
                    <div class="event-meta">
                        <div class="event-meta-item">
                            <span>📍</span> {evt_data['local']}
                        </div>
                        <div class="event-meta-item">
                            <span>🕒</span> {evt_data['hora']}
                        </div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>enter><small style='color: #888;'>Deus seja louvado</small></center>", unsafe_allow_html=True)
