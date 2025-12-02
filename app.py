import streamlit as st
import xlsxwriter
import calendar
from io import BytesIO
import datetime
from fpdf import FPDF

# ==========================================
# 1. LÓGICA DO CALENDÁRIO
# ==========================================
NOMES_MESES = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho", 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
DIAS_SEMANA_PT = ["DOMINGO", "SEGUNDA-FEIRA", "TERÇA-FEIRA", "QUARTA-FEIRA", "QUINTA-FEIRA", "SEXTA-FEIRA", "SÁBADO"]

def calcular_eventos(ano, lista_eventos):
    agenda = {}
    calendar.setfirstweekday(calendar.MONDAY)
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
                    texto = f"{evt['nome']}\n{evt['local']} AS {evt['hora']}"
                    if chave not in agenda: agenda[chave] = []
                    agenda[chave].append(texto)
    return agenda

def gerar_excel_buffer(ano, lista_eventos, uploaded_logo):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet(f"Calendário {ano}")

    # CORES
    COR_VERDE_ESCURO = '#1F4E5F'
    COR_AMARELO_NEON = '#FFFF00'
    COR_CINZA_LINHA  = '#D9D9D9'

    # FORMATOS
    fmt_ano = wb.add_format({'bold': True, 'font_size': 24, 'font_color': 'white', 'bg_color': COR_VERDE_ESCURO, 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_mes_nome = wb.add_format({'font_size': 28, 'font_color': COR_VERDE_ESCURO, 'align': 'left', 'valign': 'bottom'})
    fmt_header_sem = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': COR_VERDE_ESCURO, 'font_size': 9, 'align': 'left', 'valign': 'vcenter', 'border': 0})
    fmt_dia_box = wb.add_format({'valign': 'top', 'align': 'left', 'border': 1, 'border_color': COR_CINZA_LINHA, 'font_size': 11})
    fmt_evento_bg = wb.add_format({'valign': 'center', 'align': 'center', 'border': 1, 'border_color': COR_CINZA_LINHA, 'bg_color': COR_AMARELO_NEON, 'text_wrap': True, 'font_size': 10, 'bold': True})
    fmt_logo_celula = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})

    dados = calcular_eventos(ano, lista_eventos)
    calendar.setfirstweekday(calendar.SUNDAY)

    LINHA = 0
    for mes in range(1, 13):
        ws.write(LINHA, 0, ano, fmt_ano)
        ws.merge_range(LINHA, 1, LINHA, 5, NOMES_MESES[mes], fmt_mes_nome)
        ws.set_row(LINHA, 40)

        if uploaded_logo is not None:
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
                    if chave in dados:
                        textos_evt = "\n".join(dados[chave])
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
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.set_font("Helvetica", size=10)
    
    dados = calcular_eventos(ano, lista_eventos)
    
    for mes in range(1, 13):
        pdf.add_page()
        
        # TÍTULO DO MÊS
        pdf.set_font("Helvetica", style="B", size=18)
        pdf.set_text_color(31, 78, 95)
        mes_nome = NOMES_MESES[mes].upper().replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U").replace("Ã", "A").replace("Õ", "O").replace("Ç", "C")
        pdf.cell(0, 15, f"{mes_nome} {ano}", align="C", ln=1)
        pdf.ln(3)
        
        # CABEÇALHO DIAS DA SEMANA
        pdf.set_font("Helvetica", style="B", size=9)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(31, 78, 95)
        
        dias_semana_abr = ["SEGUNDA", "TERCA", "QUARTA", "QUINTA", "SEXTA", "SABADO", "DOMINGO"]
        largura_coluna = (pdf.w - 20) / 7
        altura_header = 7
        
        for dia_sem in dias_semana_abr:
            pdf.cell(largura_coluna, altura_header, dia_sem, border=1, align="C", fill=True)
        pdf.ln()
        
        # CALENDÁRIO EM GRADE
        calendar.setfirstweekday(calendar.MONDAY)
        cal = calendar.monthcalendar(ano, mes)
        
        pdf.set_font("Helvetica", style="B", size=12)
        pdf.set_text_color(0, 0, 0)
        altura_dia = 30
        
        for semana in cal:
            for idx, dia in enumerate(semana):
                if dia == 0:
                    # Célula vazia
                    pdf.set_fill_color(200, 200, 200)
                    pdf.cell(largura_coluna, altura_dia, "", border=1, fill=True)
                else:
                    chave = f"{ano}-{mes}-{dia}"
                    if chave in dados:
                        # Dia com evento - fundo amarelo
                        pdf.set_fill_color(255, 255, 0)
                        pdf.cell(largura_coluna, altura_dia, str(dia), border=1, align="C", fill=True)
                    else:
                        # Dia normal - fundo branco
                        pdf.set_fill_color(255, 255, 255)
                        pdf.cell(largura_coluna, altura_dia, str(dia), border=1, align="C", fill=False)
            pdf.ln()
        
        pdf.ln(8)
    
    return bytes(pdf.output())

# ==========================================
# 2. INTERFACE DO APP (STREAMLIT)
# ==========================================
st.set_page_config(page_title="Gerador CCB", page_icon="📅")
st.title("📅 Gerador de Calendário CCB")
st.write("Configure os eventos e gere sua planilha Excel ou PDF prontos.")

with st.sidebar:
    st.header("⚙️ Configuração")
    ano_escolhido = st.number_input("Ano do Calendário", value=datetime.date.today().year + 1, step=1)
    uploaded_file = st.file_uploader("Escolher Logo (Opcional)", type=['jpg', 'png'])
    logo_data = uploaded_file.getvalue() if uploaded_file else None

if 'eventos' not in st.session_state:
    st.session_state['eventos'] = [
        {"nome": "ENSAIO COM CULTO", "semana": "3", "dia_sem": "2", "interc": "Meses Ímpares", "hora": "19:30 HRS", "local": "ENTRE RIOS"},
        {"nome": "ENSAIO LOCAL", "semana": "1", "dia_sem": "5", "interc": "Todos os Meses", "hora": "19:30 HRS", "local": "SÃO PEDRO DA CIPA"},
    ]

with st.expander("➕ Adicionar Novo Evento", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        novo_nome = st.text_input("Nome", value="ENSAIO LOCAL")
        novo_dia = st.selectbox("Dia da Semana", options=[0,1,2,3,4,5,6], format_func=lambda x: ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"][x], index=5)
        novo_interc = st.selectbox("Repetição", ["Todos os Meses", "Meses Ímpares", "Meses Pares"])
    with col2:
        novo_local = st.text_input("Local", placeholder="Ex: Jaciara")
        novo_semana = st.selectbox("Semana do Mês", options=["1", "2", "3", "4", "5"], index=0)
        novo_hora = st.text_input("Hora", value="19:30 HRS")
    
    if st.button("Adicionar Evento"):
        item = {
            "nome": novo_nome.upper(),
            "local": novo_local.upper(),
            "dia_sem": str(novo_dia),
            "semana": novo_semana,
            "hora": novo_hora.upper(),
            "interc": novo_interc
        }
        st.session_state['eventos'].append(item)
        st.success("✅ Evento
