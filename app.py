import streamlit as st
import xlsxwriter
import calendar
from io import BytesIO
from datetime import datetime, date
from fpdf import FPDF

# ==========================================
# 1. LÓGICA DO CALENDÁRIO
# ==========================================
NOMES_MESES = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho", 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
DIAS_SEMANA_PT = ["DOMINGO", "SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO"]
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
st.set_page_config(page_title="Agenda CCB Jaciara", page_icon="📅", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .block-container {
        padding-top: 2rem;
        padding-bottom: 5rem;
    }
    .agenda-card {
        background-color: #ffffff;
        border-left: 5px solid #1F4E5F;
        padding: 18px;
        border-radius: 8px;
        margin-bottom: 14px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
        cursor: default;
    }
    .agenda-card:hover {
        transform: scale(1.02);
        box-shadow: 0 8px 15px rgba(0,0,0,0.15);
        border-left: 5px solid #00B4D8;
    }
    .agenda-dia {
        font-size: 28px;
        font-weight: 800;
        color: #1F4E5F;
        text-align: center;
        line-height: 1;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .agenda-sem {
        font-size: 11px;
        color: #888;
        text-align: center;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-top: 2px;
    }
    .agenda-titulo {
        font-weight: 700;
        font-size: 16px;
        color: #222;
        margin-bottom: 4px;
    }
    .agenda-local {
        color: #555;
        font-size: 14px;
        display: flex;
        align-items: center;
        gap: 5px;
    }
    .agenda-hora {
        color: #1F4E5F;
        font-weight: 600;
        font-size: 14px;
        margin-top: 4px;
        background-color: #eef6f8;
        display: inline-block;
        padding: 2px 8px;
        border-radius: 12px;
    }
    .mes-header {
        color: white;
        background: linear-gradient(90deg, #1F4E5F 0%, #2c6e85 100%);
        padding: 12px;
        border-radius: 6px;
        margin-top: 30px;
        margin-bottom: 15px;
        text-align: center;
        font-size: 18px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

st.title("📅 Ensaios Locais da Microrregião Jaciara - MT")

# --- SIDEBAR (Configurações visíveis para Admin) ---
with st.sidebar:
    st.header("Painel")
    ano_escolhido = st.number_input("Ano do Calendário", value=date.today().year + 1, step=1)
    uploaded_file = st.file_uploader("Escolher Logo (Opcional)", type=['jpg', 'png'])
    logo_data = uploaded_file.getvalue() if uploaded_file else None

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

# --- SELETOR DE MODO COM SENHA ---
modo = st.radio("Modo de Visualização:", ["🗓️ Visualizar Agenda Completa", "🔐 Configuração (Admin)"], horizontal=True)

st.divider()

if modo == "🔐 Configuração (Admin)":
    # Login Simples
    senha = st.text_input("Digite a senha de administrador:", type="password")
    
    if senha == "ccb123":  # <-- SENHA DO ADMIN AQUI
        st.success("Acesso Liberado!")
        
        with st.expander("➕ Adicionar Novo Evento", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                novo_nome = st.text_input("Nome", value="ENSAIO LOCAL")
                novo_dia = st.selectbox("Dia da Semana", options=[0,1,2,3,4,5,6], format_func=lambda x: ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"][x], index=5)
                novo_interc = st.selectbox("Repetição", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
            with col2:
                novo_local = st.text_input("Local", placeholder="Ex: Jaciara")
                novo_semana = st.selectbox("Semana do Mês", options=["1", "2", "3", "4", "5"], index=0)
                novo_hora = st.text_input("Hora", value="19:30 HRS")
            
            if st.button("Adicionar Evento"):
                item = {"nome": novo_nome.upper(), "local": novo_local.upper(), "dia_sem": str(novo_dia), "semana": novo_semana, "hora": novo_hora.upper(), "interc": novo_interc}
                st.session_state['eventos'].append(item)
                st.success("✅ Evento Adicionado!")

        st.subheader(f"📋 Lista de Eventos Cadastrados ({len(st.session_state['eventos'])})")
        for i, evt in enumerate(st.session_state['eventos']):
            dia_desc = DIAS_SEMANA_CURTO[int(evt['dia_sem'])]
            with st.container():
                col_a, col_b, col_c = st.columns([5, 2, 1])
                with col_a:
                    st.markdown(f"**{evt['nome']}**")
                    st.text(f"{evt['local']} - {evt['hora']}")
                with col_b:
                    st.info(f"{evt['semana']}ª {dia_desc} \n({evt['interc']})")
                with col_c:
                    if st.button("🗑️", key=f"del_{i}"):
                        st.session_state['eventos'].pop(i)
                        st.rerun()
            st.divider()

        st.header("🚀 Gerar Arquivos Finais")
        col_excel, col_pdf = st.columns(2)
        with col_excel:
            if st.button("📊 Gerar Excel"):
                arquivo_excel = gerar_excel_buffer(ano_escolhido, st.session_state['eventos'], logo_data)
                st.download_button(label="⬇️ BAIXAR EXCEL", data=arquivo_excel, file_name=f"Calendario_CCB_{ano_escolhido}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_pdf:
            if st.button("📄 Gerar PDF"):
                arquivo_pdf = gerar_pdf_buffer(ano_escolhido, st.session_state['eventos'])
                st.download_button(label="⬇️ BAIXAR PDF", data=arquivo_pdf, file_name=f"Calendario_CCB_{ano_escolhido}.pdf", mime="application/pdf")
    elif senha:
        st.error("Senha incorreta!")
    else:
        st.info("Por favor, digite a senha para acessar as configurações.")

else:
    # --- MODO AGENDA (PÚBLICO) ---
    # Força o seletor para o topo, e a agenda aparece por padrão
    st.header(f"🗓️ Agenda de Ensaios {ano_escolhido}")
    agenda = montar_agenda_ordenada(ano_escolhido, st.session_state['eventos'])
    
    if not agenda:
        st.warning("Nenhum evento encontrado para os critérios cadastrados.")
    else:
        mes_atual = 0
        for dt, evt_data in agenda:
            if dt.month != mes_atual:
                mes_atual = dt.month
                st.markdown(f"<div class='mes-header'>{NOMES_MESES[mes_atual]}</div>", unsafe_allow_html=True)
            
            dia_semana = DIAS_SEMANA_PT[int(dt.strftime("%w"))]
            dia_num = dt.day
            
            st.markdown(f"""
            <div class="agenda-card">
                <div style="display: flex; align-items: center;">
                    <div style="width: 90px; flex-shrink: 0; border-right: 1px solid #eee; padding-right: 15px; margin-right: 15px; text-align: center;">
                        <div class="agenda-dia">{dia_num}</div>
                        <div class="agenda-sem">{dia_semana}</div>
                    </div>
                    <div style="flex-grow: 1;">
                        <div class="agenda-titulo">{evt_data['titulo']}</div>
                        <div class="agenda-local">📍 {evt_data['local']}</div>
                        <div class="agenda-hora">🕒 {evt_data['hora']}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
