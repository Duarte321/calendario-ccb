import calendar
from datetime import date

import requests
import streamlit as st

from image_export import gerar_imagem_mes

SUPABASE_URL = "https://ovnwnzqjjjtfqjodvusi.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_uBqke5HDz9U-xSKjxhzUww_-Y0qW367"
MESES = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}


def headers():
    return {"apikey": SUPABASE_PUBLISHABLE_KEY, "Authorization": f"Bearer {SUPABASE_PUBLISHABLE_KEY}"}


def carregar_eventos():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/calendario_eventos",
        params={"select":"id,nome,local,dia_sem,semana,hora,interc","order":"id.asc"},
        headers=headers(), timeout=10,
    )
    r.raise_for_status()
    return r.json()


def carregar_avisos():
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/calendario_avisos",
        params={"select":"mes,texto","order":"mes.asc"},
        headers=headers(), timeout=10,
    )
    r.raise_for_status()
    return {int(x["mes"]): x.get("texto", "") for x in r.json()}


def calcular_agenda(ano, eventos):
    agenda = []
    calendar.setfirstweekday(calendar.SUNDAY)
    for mes in range(1, 13):
        matriz = calendar.monthcalendar(ano, mes)
        for evt in eventos:
            freq = evt["interc"]
            ativo = freq == "Todos os Meses" or (freq == "Meses Ímpares" and mes % 2) or (freq == "Meses Pares" and mes % 2 == 0)
            if not ativo:
                continue
            contador = 0
            for semana in matriz:
                numero = semana[int(evt["dia_sem"])]
                if numero:
                    contador += 1
                    if contador == int(evt["semana"]):
                        agenda.append((date(ano, mes, numero), {
                            "titulo": evt["nome"], "local": evt["local"], "hora": evt["hora"]
                        }))
                        break
    return sorted(agenda, key=lambda x: x[0])


st.set_page_config(page_title="Gerar Imagem da Agenda", page_icon="🖼️", layout="centered")
st.markdown("""
<style>
#MainMenu, footer, header {visibility:hidden}
.stApp {background:linear-gradient(180deg,#f9fafc,#f3f5f8)}
.block-container {max-width:900px;padding-top:2rem}
.hero {background:linear-gradient(135deg,#061d33,#0b2d4d);padding:25px;border-radius:20px;color:white;margin-bottom:20px;border:1px solid rgba(224,167,47,.4)}
.hero h1 {color:white;margin:0;font-size:32px}.hero p{color:#f2c45e;margin:8px 0 0}
div[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#d99a21,#f1bd50)!important;color:#061d33!important;border:none!important;font-weight:900!important;border-radius:12px!important;min-height:50px}
</style>
<div class="hero"><h1>🖼️ Gerar Imagem da Agenda</h1><p>PNG 1080 × 1350 em alta qualidade • ideal para WhatsApp e Instagram</p></div>
""", unsafe_allow_html=True)

try:
    eventos = carregar_eventos()
    avisos = carregar_avisos()
except Exception as exc:
    st.error(f"Não foi possível carregar a agenda: {exc}")
    st.stop()

c1, c2 = st.columns(2)
ano = c1.number_input("Ano", min_value=2020, max_value=2100, value=date.today().year, step=1)
mes = c2.selectbox("Mês", range(1,13), index=date.today().month-1, format_func=lambda x: MESES[x])

agenda = calcular_agenda(int(ano), eventos)

try:
    imagem = gerar_imagem_mes(int(ano), int(mes), agenda, avisos.get(int(mes), ""))
    st.subheader(f"Prévia • {MESES[int(mes)]} {int(ano)}")
    st.image(imagem.getvalue(), use_container_width=True)
    st.download_button(
        f"⬇️ BAIXAR IMAGEM DE {MESES[int(mes)].upper()}",
        data=imagem.getvalue(),
        file_name=f"Agenda_Musical_{MESES[int(mes)]}_{int(ano)}.png",
        mime="image/png",
        use_container_width=True,
    )
    st.caption("A imagem é gerada automaticamente com os dados atuais do Supabase.")
except Exception as exc:
    st.error(f"Não foi possível gerar a imagem: {exc}")
