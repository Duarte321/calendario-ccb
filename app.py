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
    # Link otimizado para mobile
    titulo = quote(f"{evt_data['titulo']} - {evt_data['local']}")
    hora_limpa = evt_data['hora'].replace("HRS", "").replace(":", "").strip()
    if len(hora_limpa) < 4: hora_limpa = "1930"
    data_inicio = f"{dt.year}{dt.month:02d}{dt.day:02d}T{hora_limpa}00"
    data_fim = f"{dt.year}{dt.month:02d}{dt.day:02d}T{int(hora_limpa[:2])+2:02d}{hora_limpa[2:]}00"
    local = quote(evt_data['local'])
    return f"https://calendar.google.com/calendar/r/eventedit?text={titulo}&dates={data_inicio}/{data_fim}&location={local}&details=Ensaio+CCB"

def gerar_excel_buffer(ano, lista_eventos, uploaded_logo):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Estilos
    header_fmt = wb.add_format({'bold': True, 'bg_color': '#1F4E5F', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    title_fmt = wb.add_format({'bold': True, 'font_size': 14, 'bg_color': '#EBF2F5', 'align': 'center', 'valign': 'vcenter'})
    data_fmt = wb.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
    center_fmt = wb.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    
    # Calcula eventos
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    
    # Por mês
    for mes in range(1, 13):
        ws = wb.add_worksheet(NOMES_MESES[mes])
        ws.set_column('A:A', 12)
        ws.set_column('B:B', 12)
        ws.set_column('C:C', 30)
        ws.set_column('D:D', 35)
        ws.set_column('E:E', 12)
        
        # Título do mês
        ws.merge_range('A1:E1', f"{NOMES_MESES[mes]} {ano}", title_fmt)
        
        # Headers
        ws.write('A2', 'Data', header_fmt)
        ws.write('B2', 'Dia Semana', header_fmt)
        ws.write('C2', 'Evento', header_fmt)
        ws.write('D2', 'Local', header_fmt)
        ws.write('E2', 'Hora', header_fmt)
        
        # Dados do mês
        linha = 3
        for dt, evt_data in agenda:
            if dt.month == mes and dt.year == ano:
                ws.write(linha-1, 0, dt.strftime('%d/%m/%Y'), data_fmt)
                ws.write(linha-1, 1, DIAS_SEMANA_CURTO[int(dt.strftime("%w"))], center_fmt)
                ws.write(linha-1, 2, evt_data['titulo'], data_fmt)
                ws.write(linha-1, 3, evt_data['local'], data_fmt)
                ws.write(linha-1, 4, evt_data['hora'], center_fmt)
                linha += 1
    
    wb.close()
    output.seek(0)
    return output

def gerar_pdf_buffer(ano, lista_eventos):
    # Configuração PDF
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Título
    pdf.set_font("Arial", "B", 20)
    pdf.set_text_color(31, 78, 95)
    pdf.cell(0, 15, f"CALENDARIO CCB JACIARA {ano}", 0, 1, 'C')
    
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)
    
    agenda = montar_agenda_ordenada(ano, lista_eventos)
    
    mes_atual = 0
    for dt, evt_data in agenda:
        if dt.month != mes_atual:
            mes_atual = dt.month
            pdf.set_font("Arial", "B", 12)
            pdf.set_text_color(31, 78, 95)
            pdf.ln(3)
            # Tratamento de acentuação manual para evitar erros de encoding
            nome_mes = NOMES_MESES[mes_atual]
            try:
                nome_mes = nome_mes.encode('latin-1', 'ignore').decode('latin-1')
            except:
                pass
            pdf.cell(0, 8, f"{nome_mes} {ano}", 0, 1, 'L')
            pdf.line(10, pdf.get_y(), 2
