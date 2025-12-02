from fpdf import FPDF

def gerar_pdf_buffer(ano, lista_eventos):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    pdf.set_font("Arial", style="B", size=24)
    pdf.cell(0, 18, f"Calendário CCB - {ano}", align="C", ln=1)
    pdf.ln(5)

    # Configuração de meses, cada mês em nova página
    for mes in range(1, 13):
        pdf.set_font("Arial", style="B", size=18)
        pdf.cell(0, 14, NOMES_MESES[mes].capitalize(), align="L", ln=1)
        pdf.set_font("Arial", style="", size=13)
        pdf.cell(0, 9, "Eventos:", ln=1)
        for evt in lista_eventos:
            desc = f"- {evt['nome']} | {evt['local']} | Semana: {evt['semana']}ª | Dia: {DIAS_SEMANA_PT[int(evt['dia_sem'])]} | Hora: {evt['hora']} | {evt['interc']}"
            pdf.cell(0, 9, desc, ln=1)
        if mes < 12:
            pdf.add_page()

    pdf_buffer = pdf.output(dest='S').encode('latin1')
    return pdf_buffer

# Botão de download no Streamlit:
if st.button("Gerar Calendário PDF"):
    pdf_data = gerar_pdf_buffer(ano_escolhido, st.session_state['eventos'])
    st.download_button(
        label="⬇️ BAIXAR PDF IMPRESSÃO",
        data=pdf_data,
        file_name=f"Calendario_CCB_{ano_escolhido}.pdf",
        mime="application/pdf"
    )
