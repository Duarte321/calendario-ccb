import calendar
from datetime import date
from io import BytesIO
from urllib.parse import quote
import requests
import streamlit as st
import xlsxwriter
from fpdf import FPDF

SUPABASE_URL = "https://ovnwnzqjjjtfqjodvusi.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_uBqke5HDz9U-xSKjxhzUww_-Y0qW367"
NOMES_MESES = {1:"JANEIRO",2:"FEVEREIRO",3:"MARÇO",4:"ABRIL",5:"MAIO",6:"JUNHO",7:"JULHO",8:"AGOSTO",9:"SETEMBRO",10:"OUTUBRO",11:"NOVEMBRO",12:"DEZEMBRO"}
DIAS_SEMANA_PT = ["DOMINGO","SEGUNDA","TERÇA","QUARTA","QUINTA","SEXTA","SÁBADO"]
DIAS_SEMANA_CURTO = ["DOM","SEG","TER","QUA","QUI","SEX","SAB"]
FREQUENCIAS = ["Todos os Meses", "Meses Ímpares", "Meses Pares"]


def _headers(key, prefer=None):
    h = {"apikey":key,"Authorization":f"Bearer {key}","Content-Type":"application/json"}
    if prefer: h["Prefer"] = prefer
    return h


def _admin_secret(): return st.secrets.get("SUPABASE_SECRET_KEY", "")
def _admin_password(): return st.secrets.get("ADMIN_PASSWORD", "")


def carregar_eventos():
    try:
        r=requests.get(f"{SUPABASE_URL}/rest/v1/calendario_eventos",params={"select":"id,nome,local,dia_sem,semana,hora,interc","order":"id.asc"},headers=_headers(SUPABASE_PUBLISHABLE_KEY),timeout=10)
        r.raise_for_status(); return r.json(), None
    except Exception as e: return [], str(e)


def carregar_avisos():
    try:
        r=requests.get(f"{SUPABASE_URL}/rest/v1/calendario_avisos",params={"select":"mes,texto","order":"mes.asc"},headers=_headers(SUPABASE_PUBLISHABLE_KEY),timeout=10)
        r.raise_for_status(); return {int(x["mes"]):x.get("texto","") for x in r.json()}, None
    except Exception as e: return {}, str(e)


def inserir_evento(evt):
    key=_admin_secret()
    if not key: raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.post(f"{SUPABASE_URL}/rest/v1/calendario_eventos",json=evt,headers=_headers(key,"return=representation"),timeout=10)
    r.raise_for_status(); return r.json()


def atualizar_evento(evento_id, evt):
    key=_admin_secret()
    if not key: raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.patch(f"{SUPABASE_URL}/rest/v1/calendario_eventos",params={"id":f"eq.{int(evento_id)}"},json=evt,headers=_headers(key,"return=representation"),timeout=10)
    r.raise_for_status()
    dados=r.json()
    if not dados: raise RuntimeError("O registro não foi encontrado para atualização.")
    return dados


def salvar_aviso(mes,texto):
    key=_admin_secret()
    if not key: raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.post(f"{SUPABASE_URL}/rest/v1/calendario_avisos",params={"on_conflict":"mes"},json={"mes":int(mes),"texto":texto},headers=_headers(key,"resolution=merge-duplicates,return=representation"),timeout=10)
    r.raise_for_status()


def excluir_aviso(mes):
    key=_admin_secret()
    if not key: raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.delete(f"{SUPABASE_URL}/rest/v1/calendario_avisos",params={"mes":f"eq.{int(mes)}"},headers=_headers(key),timeout=10)
    r.raise_for_status()


def calcular_eventos(ano,eventos):
    agenda={}; calendar.setfirstweekday(calendar.SUNDAY)
    for mes in range(1,13):
        matriz=calendar.monthcalendar(ano,mes)
        for evt in eventos:
            freq=evt["interc"]
            if not (freq=="Todos os Meses" or (freq=="Meses Ímpares" and mes%2) or (freq=="Meses Pares" and mes%2==0)): continue
            cont=0; achado=None
            for semana in matriz:
                n=semana[int(evt["dia_sem"])]
                if n:
                    cont+=1
                    if cont==int(evt["semana"]): achado=n; break
            if achado:
                agenda.setdefault(f"{ano}-{mes}-{achado}",[]).append({"titulo":evt["nome"],"local":evt["local"],"hora":evt["hora"]})
    return agenda


def montar_agenda_ordenada(ano,eventos):
    out=[]
    for chave,evts in calcular_eventos(ano,eventos).items():
        y,m,d=map(int,chave.split("-")); dt=date(y,m,d)
        out.extend((dt,e) for e in evts)
    return sorted(out,key=lambda x:x[0])


def gerar_link_google(dt,e):
    h=e["hora"].replace("HRS","").replace(":","").strip()
    if len(h)!=4 or not h.isdigit(): h="1930"
    hi=int(h[:2]); hf=min(hi+2,23)
    inicio=f"{dt.year}{dt.month:02d}{dt.day:02d}T{hi:02d}{h[2:]}00"; fim=f"{dt.year}{dt.month:02d}{dt.day:02d}T{hf:02d}{h[2:]}00"
    return f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={quote(e['titulo']+' - '+e['local'])}&dates={inicio}/{fim}&location={quote(e['local'])}&details=Ensaio+CCB&sf=true&output=xml"


def gerar_excel(ano,eventos,avisos):
    out=BytesIO(); wb=xlsxwriter.Workbook(out,{"in_memory":True}); ws=wb.add_worksheet("Calendário")
    head=wb.add_format({"bold":True,"bg_color":"#1F4E5F","font_color":"white","align":"center","border":1}); cell=wb.add_format({"border":1,"text_wrap":True,"valign":"top"}); evt_fmt=wb.add_format({"border":1,"text_wrap":True,"valign":"top","bg_color":"#FFFF00","bold":True})
    agenda=montar_agenda_ordenada(ano,eventos); mapa={}
    for dt,e in agenda: mapa.setdefault((dt.month,dt.day),[]).append(e)
    ws.set_column(0,6,18); row=0
    for mes in range(1,13):
        ws.merge_range(row,0,row,6,f"{NOMES_MESES[mes]} {ano}",head); row+=1
        for c,d in enumerate(DIAS_SEMANA_CURTO): ws.write(row,c,d,head)
        row+=1
        for semana in calendar.monthcalendar(ano,mes):
            ws.set_row(row,65)
            for c,d in enumerate(semana):
                texto="" if not d else str(d)
                fmt=cell
                if d and (mes,d) in mapa:
                    texto+= "\n"+"\n".join(f"{e['titulo']}\n{e['local']}\n{e['hora']}" for e in mapa[(mes,d)]); fmt=evt_fmt
                ws.write(row,c,texto,fmt)
            row+=1
        ws.merge_range(row,0,row,6,f"Anotações: {avisos.get(mes,'')}",cell); row+=2
    wb.close(); out.seek(0); return out


def gerar_pdf(ano,eventos,avisos):
    pdf=FPDF(); agenda=montar_agenda_ordenada(ano,eventos)
    for mes in range(1,13):
        pdf.add_page(); pdf.set_font("Arial","B",16); pdf.cell(0,10,f"{NOMES_MESES[mes]} {ano}",0,1,"C"); pdf.set_font("Arial","",9)
        for dt,e in agenda:
            if dt.month==mes: pdf.multi_cell(0,6,f"{dt.day:02d}/{mes:02d} - {e['titulo']} - {e['local']} - {e['hora']}")
        if avisos.get(mes): pdf.ln(4); pdf.set_font("Arial","B",9); pdf.multi_cell(0,6,f"AVISO: {avisos[mes]}")
    v=pdf.output(dest="S"); return v.encode("latin-1") if isinstance(v,str) else bytes(v)


st.set_page_config(page_title="Agenda CCB",page_icon="📅",layout="centered")
for k,v in {"theme":"light","nav":"Agenda","ano_base":date.today().year,"editando_id":None}.items():
    if k not in st.session_state: st.session_state[k]=v

dark=st.session_state.theme=="dark"; bg="linear-gradient(135deg,#0F2027,#203A43,#2C5364)" if dark else "linear-gradient(135deg,#F5F7FA,#C3CFE2)"; text="#fff" if dark else "#1F4E5F"; card="rgba(30,40,50,.78)" if dark else "rgba(255,255,255,.9)"
st.markdown(f"""<style>#MainMenu,footer,header{{visibility:hidden}}.stApp{{background:{bg};background-attachment:fixed}}.block-container{{max-width:780px;padding-top:1.5rem}}.title{{text-align:center;color:{text}}}.event-card{{background:{card};border-radius:18px;padding:16px;margin:12px 0;box-shadow:0 8px 28px rgba(0,0,0,.12)}}.month{{font-size:26px;font-weight:900;color:{text};margin:30px 0 8px}}.next{{background:linear-gradient(135deg,#1F4E5F,#468196);color:white;padding:20px;border-radius:20px;text-align:center;margin-bottom:24px}}.aviso{{background:#fff0f0;color:#b71c1c;border-left:4px solid #d32f2f;padding:12px;border-radius:8px}}</style>""",unsafe_allow_html=True)

_,tc=st.columns([8,1])
with tc:
    if st.button("☀️" if dark else "🌙"):
        st.session_state.theme="light" if dark else "dark"; st.rerun()
st.markdown("<h1 class='title'>Agenda CCB Jaciara</h1>",unsafe_allow_html=True)
n1,n2=st.columns(2)
if n1.button("📅 VER AGENDA",use_container_width=True): st.session_state.nav="Agenda"; st.rerun()
if n2.button("🔒 ADMIN",use_container_width=True): st.session_state.nav="Admin"; st.rerun()
st.divider()

eventos,erro_e=carregar_eventos(); avisos,erro_a=carregar_avisos()
if erro_e or erro_a: st.error("Não foi possível consultar o banco de dados agora.")

if st.session_state.nav=="Agenda":
    agenda=montar_agenda_ordenada(st.session_state.ano_base,eventos); hoje=date.today(); prox=next(((d,e) for d,e in agenda if d>=hoje),None)
    if prox:
        d,e=prox; faltam=(d-hoje).days; txt="HOJE!" if faltam==0 else f"Faltam {faltam} dias"
        st.markdown(f"<div class='next'>✨ PRÓXIMO ENSAIO • {txt}<h2>{e['titulo']}</h2>{e['local']}<br>{d:%d/%m/%Y} • {e['hora']}</div>",unsafe_allow_html=True)
    mes=0
    for d,e in agenda:
        if d.month!=mes:
            mes=d.month; st.markdown(f"<div class='month'>{NOMES_MESES[mes]} {d.year}</div>",unsafe_allow_html=True)
            if avisos.get(mes): st.markdown(f"<div class='aviso'>📢 {avisos[mes]}</div>",unsafe_allow_html=True)
        st.markdown(f"<div class='event-card'><b>{d:%d/%m} • {e['titulo']}</b><br>📍 {e['local']}<br>🕒 {e['hora']} • {DIAS_SEMANA_PT[int(d.strftime('%w'))]}<br><a href='{gerar_link_google(d,e)}' target='_blank'>🔔 Adicionar lembrete</a></div>",unsafe_allow_html=True)
else:
    st.subheader("🔒 Painel Administrativo")
    if not _admin_password() or not _admin_secret(): st.warning("Configure ADMIN_PASSWORD e SUPABASE_SECRET_KEY nos Secrets do Streamlit.")
    else:
        senha=st.text_input("Senha de Acesso",type="password")
        if senha==_admin_password():
            st.success("✅ Acesso liberado"); st.session_state.ano_base=st.number_input("Ano de Referência",2020,2100,int(st.session_state.ano_base))
            abas=st.tabs(["➕ Novo Evento","📝 Avisos","✏️ Editar Eventos","📥 Downloads"])
            with abas[0]:
                with st.form("novo",clear_on_submit=True):
                    nome=st.text_input("Nome","ENSAIO LOCAL"); local=st.text_input("Local"); dia=st.selectbox("Dia",range(7),format_func=lambda x:DIAS_SEMANA_PT[x]); semana=st.selectbox("Semana",[1,2,3,4,5]); hora=st.text_input("Hora","19:30 HRS"); freq=st.selectbox("Frequência",FREQUENCIAS)
                    if st.form_submit_button("💾 Salvar Evento",use_container_width=True):
                        if not local.strip(): st.error("Informe o local.")
                        else:
                            try: inserir_evento({"nome":nome.strip().upper(),"local":local.strip().upper(),"dia_sem":dia,"semana":semana,"hora":hora.strip().upper(),"interc":freq}); st.rerun()
                            except Exception as e: st.error(f"Erro: {e}")
            with abas[1]:
                mes=st.selectbox("Mês",range(1,13),format_func=lambda x:NOMES_MESES[x]); aviso=st.text_area("Aviso",value=avisos.get(mes,"")); a,b=st.columns(2)
                if a.button("💾 Salvar Aviso",use_container_width=True):
                    try: salvar_aviso(mes,aviso); st.rerun()
                    except Exception as e: st.error(f"Erro: {e}")
                if b.button("🗑️ Apagar Aviso",use_container_width=True):
                    try: excluir_aviso(mes); st.rerun()
                    except Exception as e: st.error(f"Erro: {e}")
            with abas[2]:
                if not eventos: st.info("Nenhum evento cadastrado.")
                opcoes={e["id"]:f"{e['local']} — {e['semana']}ª {DIAS_SEMANA_CURTO[int(e['dia_sem'])]} — {e['hora']}" for e in eventos}
                if opcoes:
                    eid=st.selectbox("Selecione o evento para editar",list(opcoes),format_func=lambda x:opcoes[x]); evt=next(e for e in eventos if e["id"]==eid)
                    with st.form(f"editar_{eid}"):
                        enome=st.text_input("Nome",evt["nome"]); elocal=st.text_input("Local",evt["local"]); edia=st.selectbox("Dia",range(7),index=int(evt["dia_sem"]),format_func=lambda x:DIAS_SEMANA_PT[x]); esem=st.selectbox("Semana",[1,2,3,4,5],index=int(evt["semana"])-1); ehora=st.text_input("Hora",evt["hora"]); efreq=st.selectbox("Frequência",FREQUENCIAS,index=FREQUENCIAS.index(evt["interc"]))
                        if st.form_submit_button("✅ SALVAR ALTERAÇÕES",use_container_width=True):
                            if not elocal.strip(): st.error("Informe o local.")
                            else:
                                try: atualizar_evento(eid,{"nome":enome.strip().upper(),"local":elocal.strip().upper(),"dia_sem":edia,"semana":esem,"hora":ehora.strip().upper(),"interc":efreq}); st.success("Evento atualizado no Supabase!"); st.rerun()
                                except Exception as e: st.error(f"Erro ao atualizar: {e}")
            with abas[3]:
                st.download_button("⬇️ Baixar Excel",gerar_excel(st.session_state.ano_base,eventos,avisos),f"Calendario_{st.session_state.ano_base}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                st.download_button("⬇️ Baixar PDF",gerar_pdf(st.session_state.ano_base,eventos,avisos),f"Calendario_{st.session_state.ano_base}.pdf",mime="application/pdf",use_container_width=True)
        elif senha: st.error("❌ Senha incorreta")
