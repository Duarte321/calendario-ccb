import streamlit as st
import xlsxwriter
import calendar
import tempfile
import os
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

# ===== FUNÇÕES DE ARQUIVO ATUALIZADAS =====
def gerar_excel_todos_meses(ano, lista_eventos, avisos, logo_bytes=None):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Calendário")
    
    # Formatações
    header_mes = wb.add_format({'bold': True, 'font_size': 14, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    header_dias = wb.add_format({'bold': True, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 10})
    cell_dia = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 'bold': True})
    
    # ATUALIZAÇÃO: Font Size alterado para 10 conforme solicitado
    cell_evento = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 'text_wrap': True, 'bg_color': '#FFFF00', 'bold': True})
    
    cell_aviso = wb.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'font_size': 10, 'text_wrap': True, 'bg_color': '#FFCDD2', 'bold': True, 'font_color': '#B71C1C'})
    cell_vazio = wb.add_format({'border': 1, 'bg_color': '#E0E0E0'})
    
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    eventos_dict = {}
    for dt, evt_data in agenda:
        chave = f"{dt.year}-{dt.month}-{dt.day}"
        if chave not in eventos_dict: eventos_dict[chave] = []
        eventos_dict[chave].append(evt_data)
        
    for col in range(7): ws.set_column(col, col, 20) # Aumentei levemente a largura
    
    current_row = 0
    
    # Inserir Logo no topo se existir
    if logo_bytes:
        logo_bytes.seek(0)
        # Insere flutuante sobre a célula A1 (com deslocamento visual) ou em um cabeçalho dedicado
        # Aqui, vamos inserir no topo da planilha, antes dos meses, ou repetir a cada mês se preferir.
        # Para simplificar, inserimos apenas uma vez no topo ou usamos offset.
        # Opção: Inserir na célula A1 com escala ajustada.
        ws.insert_image('A1', 'logo.png', {'image_data': logo_bytes, 'x_scale': 0.15, 'y_scale': 0.15, 'x_offset': 10, 'y_offset': 10})

    for mes in range(1, 13):
        nome_mes = NOMES_MESES[mes]
        ws.merge_range(current_row, 0, current_row, 6, f"{nome_mes} {ano}", header_mes)
        ws.set_row(current_row, 35)
        current_row += 1
        
        for col, dia in enumerate(DIAS_SEMANA_CURTO):
            ws.write(current_row, col, dia, header_dias)
        ws.set_row(current_row, 20)
        current_row += 1
        
        cal_matrix = calendar.monthcalendar(ano, mes)
        for semana in cal_matrix:
            ws.set_row(current_row, 85) # Altura aumentada para acomodar fonte maior
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

def gerar_pdf_calendario(ano, lista_eventos, avisos, logo_bytes=None):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    eventos_dict = {}
    
    # Tratar Logo Temp
    logo_path = None
    if logo_bytes:
        logo_bytes.seek(0)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
            tmp.write(logo_bytes.read())
            logo_path = tmp.name

    for dt, evt_data in agenda:
        chave = f"{dt.year}-{dt.month}-{dt.day}"
        if chave not in eventos_dict: eventos_dict[chave] = []
        eventos_dict[chave].append(evt_data)
        
    for mes in range(1, 13):
        pdf.add_page()
        
        # LOGO NO PDF
        if logo_path:
            try:
                pdf.image(logo_path, x=10, y=8, h=18) # Ajuste x, y, h conforme necessário
            except:
                pass

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
                        pdf.set_font("Arial", "B", 7) # PDF um pouco menor para caber, Excel usei 10
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

    # Limpeza Temp
    if logo_path and os.path.exists(logo_path):
        os.remove(logo_path)
        
    try:
        val = pdf.output(dest='S')
        if isinstance(val, str): return val.encode('latin-1')
        return bytes(val)
    except: return bytes(pdf.output())

# ==========================================
# 2. VISUAL & CONFIG
# ==========================================
st.set_page_config(page_title="Agenda CCB", page_icon="📅", layout="centered", initial_sidebar_state="collapsed")

if 'theme' not in st.session_state: st.session_state['theme'] = 'light'
# ... (MANTIDA A CONFIGURAÇÃO DE CSS ORIGINAL) ...

# (
