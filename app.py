import calendar
import html
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
NOMES_MESES_TITULO = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
DIAS_SEMANA_PT=["DOMINGO","SEGUNDA","TERÇA","QUARTA","QUINTA","SEXTA","SÁBADO"]
DIAS_SEMANA_CURTO=["DOM","SEG","TER","QUA","QUI","SEX","SÁB"]
FREQUENCIAS=["Todos os Meses","Meses Ímpares","Meses Pares"]


def _headers(key,prefer=None):
    h={"apikey":key,"Authorization":f"Bearer {key}","Content-Type":"application/json"}
    if prefer:h["Prefer"]=prefer
    return h

def _admin_secret():return st.secrets.get("SUPABASE_SECRET_KEY","")
def _admin_password():return st.secrets.get("ADMIN_PASSWORD","")
def definir_notificacao(m):st.session_state["flash_message"]=m

def mostrar_notificacao():
    m=st.session_state.get("flash_message")
    if m:
        st.success(m)
        try:st.toast(m,icon="✅")
        except Exception:pass
        st.session_state["flash_message"]=None

def carregar_eventos():
    try:
        r=requests.get(f"{SUPABASE_URL}/rest/v1/calendario_eventos",params={"select":"id,nome,local,dia_sem,semana,hora,interc","order":"id.asc"},headers=_headers(SUPABASE_PUBLISHABLE_KEY),timeout=10);r.raise_for_status();return r.json(),None
    except Exception as e:return [],str(e)

def carregar_avisos():
    try:
        r=requests.get(f"{SUPABASE_URL}/rest/v1/calendario_avisos",params={"select":"mes,texto","order":"mes.asc"},headers=_headers(SUPABASE_PUBLISHABLE_KEY),timeout=10);r.raise_for_status();return {int(x["mes"]):x.get("texto","") for x in r.json()},None
    except Exception as e:return {},str(e)

def inserir_evento(evt):
    k=_admin_secret()
    if not k:raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.post(f"{SUPABASE_URL}/rest/v1/calendario_eventos",json=evt,headers=_headers(k,"return=representation"),timeout=10);r.raise_for_status();return r.json()

def atualizar_evento(i,evt):
    k=_admin_secret()
    if not k:raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.patch(f"{SUPABASE_URL}/rest/v1/calendario_eventos",params={"id":f"eq.{int(i)}"},json=evt,headers=_headers(k,"return=representation"),timeout=10);r.raise_for_status();d=r.json()
    if not d:raise RuntimeError("O registro não foi encontrado para atualização.")
    return d

def excluir_evento(i):
    k=_admin_secret()
    if not k:raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.delete(f"{SUPABASE_URL}/rest/v1/calendario_eventos",params={"id":f"eq.{int(i)}"},headers=_headers(k,"return=representation"),timeout=10);r.raise_for_status();return r.json()

def salvar_aviso(mes,texto):
    k=_admin_secret()
    if not k:raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.post(f"{SUPABASE_URL}/rest/v1/calendario_avisos",params={"on_conflict":"mes"},json={"mes":int(mes),"texto":texto},headers=_headers(k,"resolution=merge-duplicates,return=representation"),timeout=10);r.raise_for_status()

def excluir_aviso(mes):
    k=_admin_secret()
    if not k:raise RuntimeError("SUPABASE_SECRET_KEY não configurada.")
    r=requests.delete(f"{SUPABASE_URL}/rest/v1/calendario_avisos",params={"mes":f"eq.{int(mes)}"},headers=_headers(k),timeout=10);r.raise_for_status()


def calcular_eventos(ano,eventos):
    agenda={};calendar.setfirstweekday(calendar.SUNDAY)
    for mes in range(1,13):
        matriz=calendar.monthcalendar(ano,mes)
        for evt in eventos:
            f=evt["interc"]
            if not(f=="Todos os Meses" or(f=="Meses Ímpares" and mes%2!=0)or(f=="Meses Pares" and mes%2==0)):continue
            c=0;achado=None
            for semana in matriz:
                n=semana[int(evt["dia_sem"])]
                if n:
                    c+=1
                    if c==int(evt["semana"]):achado=n;break
            if achado:agenda.setdefault(f"{ano}-{mes}-{achado}",[]).append({"titulo":evt["nome"],"local":evt["local"],"hora":evt["hora"]})
    return agenda

def montar_agenda_ordenada(ano,eventos):
    out=[]
    for chave,evts in calcular_eventos(ano,eventos).items():
        y,m,d=map(int,chave.split("-"));dt=date(y,m,d);out.extend((dt,e)for e in evts)
    return sorted(out,key=lambda x:x[0])

def gerar_link_google(dt,e):
    h=e["hora"].replace("HRS","").replace(":","").strip()
    if len(h)!=4 or not h.isdigit():h="1930"
    hi=int(h[:2]);hf=min(hi+2,23);ini=f"{dt.year}{dt.month:02d}{dt.day:02d}T{hi:02d}{h[2:]}00";fim=f"{dt.year}{dt.month:02d}{dt.day:02d}T{hf:02d}{h[2:]}00"
    return f"https://calendar.google.com/calendar/render?action=TEMPLATE&text={quote(e['titulo']+' - '+e['local'])}&dates={ini}/{fim}&location={quote(e['local'])}&details=Ensaio+CCB&sf=true&output=xml"


def gerar_excel(ano,eventos,avisos):
    out=BytesIO();wb=xlsxwriter.Workbook(out,{"in_memory":True});ws=wb.add_worksheet("Calendário")
    head=wb.add_format({"bold":True,"bg_color":"#0B2D4D","font_color":"white","align":"center","border":1});cell=wb.add_format({"border":1,"text_wrap":True,"valign":"top"});evt_fmt=wb.add_format({"border":1,"text_wrap":True,"valign":"top","bg_color":"#FFF1C6","bold":True})
    agenda=montar_agenda_ordenada(ano,eventos);mapa={}
    for dt,e in agenda:mapa.setdefault((dt.month,dt.day),[]).append(e)
    ws.set_column(0,6,18);row=0
    for mes in range(1,13):
        ws.merge_range(row,0,row,6,f"{NOMES_MESES[mes]} {ano}",head);row+=1
        for c,d in enumerate(DIAS_SEMANA_CURTO):ws.write(row,c,d,head)
        row+=1
        for semana in calendar.monthcalendar(ano,mes):
            ws.set_row(row,65)
            for c,d in enumerate(semana):
                t="" if not d else str(d);fmt=cell
                if d and(mes,d)in mapa:t+="\n"+"\n".join(f"{e['titulo']}\n{e['local']}\n{e['hora']}"for e in mapa[(mes,d)]);fmt=evt_fmt
                ws.write(row,c,t,fmt)
            row+=1
        ws.merge_range(row,0,row,6,f"Anotações: {avisos.get(mes,'')}",cell);row+=2
    wb.close();out.seek(0);return out

def texto_pdf_seguro(v):
    return str(v or"").replace("—","-").replace("–","-").replace("•","-").encode("latin-1","replace").decode("latin-1")

def gerar_pdf(ano,eventos,avisos):
    pdf=FPDF(orientation="P",unit="mm",format="A4");pdf.set_auto_page_break(auto=False);agenda=montar_agenda_ordenada(ano,eventos);ed={}
    for dt,e in agenda:ed.setdefault((dt.month,dt.day),[]).append(e)
    margem=10;lt=190;lc=lt/7;hc=7;hl=32
    for mes in range(1,13):
        pdf.add_page();pdf.set_fill_color(11,45,77);pdf.rect(margem,10,lt,15,"F");pdf.set_xy(margem,10);pdf.set_font("Arial","B",16);pdf.set_text_color(255,255,255);pdf.cell(lt,15,texto_pdf_seguro(f"{NOMES_MESES[mes]} {ano}"),0,0,"C")
        yh=30;pdf.set_xy(margem,yh);pdf.set_font("Arial","B",8);pdf.set_fill_color(11,45,77)
        for ds in DIAS_SEMANA_CURTO:pdf.cell(lc,hc,ds,1,0,"C",fill=True)
        matriz=calendar.monthcalendar(ano,mes);topo=yh+hc
        for li,semana in enumerate(matriz):
            y=topo+li*hl
            for ci,d in enumerate(semana):
                x=margem+ci*lc;tem=d!=0 and(mes,d)in ed
                pdf.set_fill_color(*((235,235,235) if d==0 else (255,241,198) if tem else (255,255,255)));pdf.set_draw_color(0,0,0);pdf.rect(x,y,lc,hl,"DF")
                if not d:continue
                pdf.set_xy(x+1.5,y+1.3);pdf.set_font("Arial","B",10);pdf.set_text_color(0,0,0);pdf.cell(7,4.5,str(d),0,0,"L")
                if tem:
                    cy=y+7
                    for e in ed[(mes,d)]:
                        pdf.set_xy(x+1.5,cy);pdf.set_font("Arial","B",6.3);pdf.multi_cell(lc-3,3.0,texto_pdf_seguro(f"{e['titulo']}\n{e['local']}\n{e['hora']}"),0,"L");cy=pdf.get_y()+.8
                        if cy>y+hl-2:break
        ya=topo+len(matriz)*hl+4;av=texto_pdf_seguro(avisos.get(mes,""));pdf.set_xy(margem,ya);pdf.set_text_color(0,0,0);pdf.set_font("Arial","B",9)
        if av:
            pdf.set_fill_color(255,240,240);pdf.cell(lt,6,"Anotacoes / Avisos Importantes:",1,1,"L",fill=True);pdf.set_x(margem);pdf.set_text_color(160,0,0);pdf.multi_cell(lt,6,av,1,"L",fill=True)
        else:
            pdf.set_fill_color(255,255,255);pdf.cell(lt,6,"Anotacoes:",1,1,"L",fill=True);pdf.set_x(margem);pdf.cell(lt,10,"",1,0,"L",fill=True)
    v=pdf.output(dest="S");return v.encode("latin-1")if isinstance(v,str)else bytes(v)


st.set_page_config(page_title="Agenda Musical | Região de Jaciara",page_icon="🎼",layout="wide")
for k,v in {"nav":"Agenda","ano_base":date.today().year,"mes_visual":date.today().month,"flash_message":None}.items():
    if k not in st.session_state:st.session_state[k]=v

st.markdown('''<style>
:root{--navy:#061d33;--navy2:#0b2d4d;--gold:#e0a72f;--gold2:#f2c45e;--ink:#13243a;--muted:#667085;--line:#e5e7eb}
#MainMenu,footer,header{visibility:hidden}.stApp{background:radial-gradient(circle at 90% 5%,rgba(224,167,47,.08),transparent 24%),linear-gradient(180deg,#f9fafc,#f3f5f8);color:var(--ink)}.block-container{max-width:1500px;padding-top:0!important;padding-left:1.4rem;padding-right:1.4rem;padding-bottom:2rem}.premium-header{margin:0 -2rem 24px;padding:24px 3rem;background:linear-gradient(115deg,#04192c,#082844 58%,#061d33);color:white;box-shadow:0 8px 25px rgba(5,29,51,.18);border-bottom:1px solid rgba(224,167,47,.35)}.brand-row{display:flex;align-items:center;justify-content:space-between;gap:22px;flex-wrap:wrap}.brand-wrap{display:flex;align-items:center;gap:16px}.brand-icon{width:54px;height:54px;border:1px solid rgba(224,167,47,.55);border-radius:15px;display:flex;align-items:center;justify-content:center;font-size:31px;color:var(--gold2);background:rgba(255,255,255,.03)}.brand-title{font-family:Georgia,serif;font-size:34px;font-weight:700;line-height:1;color:white}.brand-sub{margin-top:7px;color:var(--gold2);font-size:15px;font-weight:600}.header-note{color:#dbe4ee;font-size:13px;text-align:right;line-height:1.5}.section-title{display:flex;align-items:center;gap:10px;font-size:18px;font-weight:850;color:var(--navy2);margin:18px 0 14px}.next-card{background:radial-gradient(circle at 85% 30%,rgba(224,167,47,.13),transparent 30%),linear-gradient(150deg,#082844,#061d33);color:white;border:1px solid rgba(224,167,47,.28);border-radius:18px;padding:22px;box-shadow:0 12px 28px rgba(6,29,51,.18);min-height:300px}.next-kicker{color:var(--gold2);font-size:12px;font-weight:900;letter-spacing:.9px}.next-title{font-size:28px;font-weight:850;margin:10px 0 4px;line-height:1.1}.next-local{color:#dfe7ef;font-size:15px;margin-bottom:18px}.next-rule{height:1px;background:rgba(255,255,255,.13);margin:17px 0}.next-meta{display:flex;gap:12px;align-items:flex-start;margin:14px 0}.next-meta .sym{font-size:20px;color:var(--gold2);width:25px}.next-meta b{display:block;font-size:15px;color:white}.next-meta span{color:#c7d2df;font-size:12px}.notice-card,.calendar-shell,.admin-shell{background:white;border:1px solid var(--line);border-radius:16px;box-shadow:0 8px 24px rgba(16,24,40,.07)}.notice-card{padding:16px;margin-top:16px}.notice-head{font-weight:850;color:var(--navy2);font-size:16px}.notice-body{margin-top:12px;color:#35445a;font-size:14px;line-height:1.55}.notice-pill{display:inline-block;margin-top:12px;background:#fff1c6;color:#8a5d00;border:1px solid #f0d58a;border-radius:999px;padding:5px 10px;font-size:11px;font-weight:800}.info-card,.share-card{background:linear-gradient(150deg,#082844,#061d33);color:white;padding:20px;border:1px solid rgba(224,167,47,.22);border-radius:16px;box-shadow:0 8px 24px rgba(16,24,40,.12)}.info-card h3,.share-card h3{color:var(--gold2);font-size:16px;margin:0 0 14px}.info-card p,.share-card p{color:#e0e7ef;font-size:14px;line-height:1.65}.gold-line{height:1px;background:rgba(224,167,47,.5);margin:18px 0}.calendar-shell{padding:18px}.calendar-title{font-size:21px;font-weight:900;color:var(--navy2)}.cal-scroll{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch;padding-bottom:6px;scrollbar-width:thin}.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);border-left:1px solid #d9dee6;border-top:1px solid #d9dee6;overflow:hidden;border-radius:10px}.cal-head{background:linear-gradient(180deg,#0b2d4d,#082844);color:white;text-align:center;font-size:12px;font-weight:850;padding:11px 4px;border-right:1px solid rgba(255,255,255,.18);border-bottom:1px solid #d9dee6}.cal-cell{min-height:115px;background:white;padding:9px;border-right:1px solid #d9dee6;border-bottom:1px solid #d9dee6}.cal-cell.empty{background:#f3f4f6}.cal-cell.event{background:linear-gradient(145deg,#fffaf0,#fff0bf)}.day-number{font-weight:900;color:#1b2738;font-size:15px;margin-bottom:7px}.cell-event{font-size:11px;line-height:1.35;color:#1b2738;margin-top:5px;font-weight:650}.cell-event .time{font-size:10px;color:#6a4c0c;margin-top:4px;font-weight:800}.mobile-swipe{display:none;color:#667085;font-size:12px;font-weight:700;text-align:center;margin:7px 0}.legend{display:flex;align-items:center;gap:8px;font-size:12px;color:#556274;margin:12px 4px 2px}.legend-box{width:16px;height:16px;border-radius:5px;background:#fff0bf;border:1px solid #f1d783}.month-divider{margin:30px 0 12px;display:flex;align-items:center;gap:12px;color:var(--navy2);font-size:21px;font-weight:900}.month-divider:after{content:"";height:1px;background:#dfe3e9;flex:1}.event-list-card{background:white;border:1px solid #e4e7ec;border-left:4px solid var(--gold);border-radius:12px;padding:14px 16px;margin:9px 0;box-shadow:0 5px 15px rgba(16,24,40,.04)}.event-list-title{font-weight:850;color:var(--navy2);font-size:15px}.event-list-info{color:#667085;font-size:13px;margin-top:4px}.event-list-card a{display:inline-block;margin-top:8px;color:#8a5d00;text-decoration:none;font-size:12px;font-weight:800}.footer-premium{margin:30px -2rem -2rem;background:linear-gradient(115deg,#04192c,#082844);color:#dbe4ee;padding:24px 3rem;border-top:1px solid rgba(224,167,47,.35)}.footer-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:24px}.footer-title{color:var(--gold2);font-weight:850;font-size:13px;margin-bottom:7px}.footer-text{font-size:12px;line-height:1.6}.footer-bottom{text-align:center;margin-top:18px;padding-top:14px;border-top:1px solid rgba(255,255,255,.12);color:var(--gold2);font-size:12px}div[data-testid="stButton"]>button{border-radius:10px!important;font-weight:800!important;min-height:42px}div[data-testid="stButton"]>button[kind="primary"]{background:linear-gradient(135deg,#d99a21,#f1bd50)!important;color:#081c30!important;border:none!important}div[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#0b2d4d,#061d33)!important;color:white!important;border:1px solid rgba(224,167,47,.35)!important;border-radius:10px!important;font-weight:800!important}div[data-baseweb="select"]>div,input,textarea{border-radius:10px!important}.stTabs [data-baseweb="tab-list"]{gap:8px}.stTabs [data-baseweb="tab"]{border-radius:10px 10px 0 0;padding:8px 14px;font-weight:750}@media(max-width:900px){.premium-header{margin:0 -.8rem 16px;padding:18px 1rem}.brand-title{font-size:27px}.brand-icon{width:48px;height:48px;font-size:27px}.header-note{display:none}.block-container{padding-left:.45rem;padding-right:.45rem}.calendar-shell{padding:12px}.calendar-title{font-size:19px;margin-bottom:5px}.cal-scroll{margin:0 -2px;width:calc(100% + 4px)}.cal-grid{min-width:760px;grid-template-columns:repeat(7,108px)}.cal-head{font-size:13px;padding:12px 5px}.cal-cell{min-height:118px;padding:8px}.day-number{font-size:18px}.cell-event{font-size:12px;line-height:1.4}.cell-event .time{font-size:11px}.mobile-swipe{display:block}.next-card{min-height:auto}.next-title{font-size:25px}.footer-grid{grid-template-columns:1fr}.footer-premium{margin:25px -.8rem -2rem;padding:22px 1.2rem}.event-list-title{font-size:16px}.event-list-info{font-size:14px}}
</style>''',unsafe_allow_html=True)

st.markdown('''<div class="premium-header"><div class="brand-row"><div class="brand-wrap"><div class="brand-icon">𝄞</div><div><div class="brand-title">Agenda Musical</div><div class="brand-sub">Região de Jaciara - MT</div></div></div><div class="header-note">Ensaios Locais • Avisos • Organização Musical<br><span style="color:#f2c45e;font-weight:700">Calendário oficial da região</span></div></div></div>''',unsafe_allow_html=True)

na,nb,nc=st.columns([2,2,7])
with na:
    if st.button("🏠 INÍCIO / AGENDA",use_container_width=True,type="primary" if st.session_state.nav=="Agenda" else "secondary"):st.session_state.nav="Agenda";st.rerun()
with nb:
    if st.button("🔐 ÁREA ADMIN",use_container_width=True,type="primary" if st.session_state.nav=="Admin" else "secondary"):st.session_state.nav="Admin";st.rerun()

eventos,erro_e=carregar_eventos();avisos,erro_a=carregar_avisos()
if erro_e or erro_a:st.error("Não foi possível consultar o banco de dados agora.")

if st.session_state.nav=="Agenda":
    agenda=montar_agenda_ordenada(st.session_state.ano_base,eventos);hoje=date.today();prox=next(((d,e)for d,e in agenda if d>=hoje),None)
    left,center,right=st.columns([1.05,2.65,1.0],gap="medium")
    with left:
        if prox:
            d,e=prox;faltam=(d-hoje).days;status="HOJE!" if faltam==0 else f"Faltam {faltam} dias"
            st.markdown(f'''<div class="next-card"><div class="next-kicker">PRÓXIMO ENSAIO • {html.escape(status)}</div><div class="next-title">{html.escape(e['titulo'].title())}</div><div class="next-local">📍 {html.escape(e['local'].title())}</div><div class="next-rule"></div><div class="next-meta"><div class="sym">📅</div><div><b>{d.strftime('%d/%m/%Y')}</b><span>{DIAS_SEMANA_PT[int(d.strftime('%w'))].title()}</span></div></div><div class="next-meta"><div class="sym">🕒</div><div><b>{html.escape(e['hora'])}</b><span>Horário</span></div></div><div class="next-meta"><div class="sym">🎼</div><div><b>Ensaio Local</b><span>Agenda musical</span></div></div></div>''',unsafe_allow_html=True)
        ma=st.session_state.mes_visual;av=avisos.get(ma,"");st.markdown('<div class="notice-card"><div class="notice-head">🔔 AVISOS IMPORTANTES</div>',unsafe_allow_html=True)
        st.markdown(f'<div class="notice-body">{html.escape(av) if av else "Nenhum aviso cadastrado para este mês."}</div>'+ (f'<span class="notice-pill">{NOMES_MESES[ma]}</span>' if av else '')+'</div>',unsafe_allow_html=True)
    with center:
        st.markdown('<div class="calendar-shell">',unsafe_allow_html=True);ct1,ct2=st.columns([3,1.15])
        with ct1:st.markdown(f'<div class="calendar-title">📅 CALENDÁRIO {st.session_state.ano_base}</div>',unsafe_allow_html=True)
        with ct2:
            mes=st.selectbox("Mês",range(1,13),index=int(st.session_state.mes_visual)-1,format_func=lambda x:NOMES_MESES_TITULO[x],label_visibility="collapsed");st.session_state.mes_visual=mes
        ad={}
        for d,e in agenda:
            if d.month==mes:ad.setdefault(d.day,[]).append(e)
        p=['<div class="mobile-swipe">↔ Deslize para ver todos os dias</div><div class="cal-scroll"><div class="cal-grid">']+[f'<div class="cal-head">{x}</div>' for x in DIAS_SEMANA_CURTO]
        for semana in calendar.monthcalendar(st.session_state.ano_base,mes):
            for dia in semana:
                if dia==0:p.append('<div class="cal-cell empty"></div>');continue
                p.append(f'<div class="{"cal-cell event" if dia in ad else "cal-cell"}"><div class="day-number">{dia}</div>')
                for e in ad.get(dia,[]):p.append(f'<div class="cell-event">{html.escape(e["titulo"].title())}<br>{html.escape(e["local"].title())}<div class="time">◷ {html.escape(e["hora"])}</div></div>')
                p.append('</div>')
        p.append('</div></div><div class="legend"><span class="legend-box"></span>Dias com ensaio</div>');st.markdown("".join(p),unsafe_allow_html=True);st.markdown('</div>',unsafe_allow_html=True)
    with right:
        st.markdown('''<div class="info-card"><h3>🎵 SOBRE A AGENDA</h3><p>Agenda dos Ensaios Locais e Avisos da Região de Jaciara - MT.</p><div class="gold-line"></div><p><b style="color:#f2c45e">Juntos em harmonia</b><br>para a glória de Deus!</p></div><div class="share-card" style="margin-top:16px"><h3>↗ COMPARTILHE</h3><p>Compartilhe a agenda com os irmãos e mantenha todos informados.</p></div>''',unsafe_allow_html=True)
    st.markdown('<div class="section-title">🎼 PRÓXIMOS ENSAIOS</div>',unsafe_allow_html=True);mm=0
    for d,e in agenda:
        if d<hoje:continue
        if d.month!=mm:
            mm=d.month;st.markdown(f'<div class="month-divider">{NOMES_MESES_TITULO[mm]} {d.year}</div>',unsafe_allow_html=True)
            if avisos.get(mm):st.info(f"📢 {avisos[mm]}")
        st.markdown(f'<div class="event-list-card"><div class="event-list-title">{d.strftime("%d/%m")} • {html.escape(e["titulo"].title())}</div><div class="event-list-info">📍 {html.escape(e["local"].title())} &nbsp;&nbsp; 🕒 {html.escape(e["hora"])} • {DIAS_SEMANA_PT[int(d.strftime("%w"))].title()}</div><a href="{gerar_link_google(d,e)}" target="_blank">🔔 ADICIONAR AO GOOGLE CALENDAR</a></div>',unsafe_allow_html=True)
    st.markdown('''<div class="footer-premium"><div class="footer-grid"><div><div class="footer-title">📅 ORGANIZE-SE</div><div class="footer-text">Planeje sua participação e acompanhe ensaios e avisos importantes.</div></div><div><div class="footer-title">💬 DÚVIDAS OU SUGESTÕES</div><div class="footer-text">Entre em contato com a administração da sua localidade.</div></div><div><div class="footer-title">♥ FEITO COM DEDICAÇÃO</div><div class="footer-text">Uma agenda simples e organizada para servir à irmandade.</div></div></div><div class="footer-bottom">Agenda Musical • Região de Jaciara - MT</div></div>''',unsafe_allow_html=True)
else:
    st.markdown('<div class="admin-shell" style="padding:22px;margin-top:16px"><div class="section-title">🔐 PAINEL ADMINISTRATIVO</div>',unsafe_allow_html=True)
    if not _admin_password() or not _admin_secret():st.warning("Configure ADMIN_PASSWORD e SUPABASE_SECRET_KEY nos Secrets do Streamlit.")
    else:
        senha=st.text_input("Senha de Acesso",type="password")
        if senha==_admin_password():
            st.success("✅ Acesso liberado");mostrar_notificacao();st.session_state.ano_base=st.number_input("Ano de Referência",2020,2100,int(st.session_state.ano_base),step=1);abas=st.tabs(["➕ Novo Evento","📝 Avisos","✏️ Gerenciar Eventos","📥 Downloads"])
            with abas[0]:
                with st.form("novo",clear_on_submit=True):
                    nome=st.text_input("Nome","ENSAIO LOCAL");local=st.text_input("Local");dia=st.selectbox("Dia",range(7),format_func=lambda x:DIAS_SEMANA_PT[x]);semana=st.selectbox("Semana",[1,2,3,4,5]);hora=st.text_input("Hora","19:30 HRS");freq=st.selectbox("Frequência",FREQUENCIAS)
                    if st.form_submit_button("💾 Salvar Evento",use_container_width=True):
                        if not local.strip():st.error("Informe o local.")
                        else:
                            try:inserir_evento({"nome":nome.strip().upper(),"local":local.strip().upper(),"dia_sem":dia,"semana":semana,"hora":hora.strip().upper(),"interc":freq});definir_notificacao("✅ Evento salvo com sucesso!");st.rerun()
                            except Exception as e:st.error(f"Erro ao salvar: {e}")
            with abas[1]:
                mes=st.selectbox("Mês",range(1,13),format_func=lambda x:NOMES_MESES[x]);av=st.text_area("Aviso",value=avisos.get(mes,""));c1,c2=st.columns(2)
                if c1.button("💾 Salvar Aviso",use_container_width=True):
                    try:salvar_aviso(mes,av);definir_notificacao("✅ Aviso salvo com sucesso!");st.rerun()
                    except Exception as e:st.error(f"Erro: {e}")
                if c2.button("🗑️ Apagar Aviso",use_container_width=True):
                    try:excluir_aviso(mes);definir_notificacao("✅ Aviso excluído com sucesso!");st.rerun()
                    except Exception as e:st.error(f"Erro: {e}")
            with abas[2]:
                if not eventos:st.info("Nenhum evento cadastrado.")
                else:
                    op={e["id"]:f"{e['local']} — {e['semana']}ª {DIAS_SEMANA_CURTO[int(e['dia_sem'])]} — {e['hora']}" for e in eventos};eid=st.selectbox("Selecione o evento",list(op),format_func=lambda x:op[x]);evt=next(e for e in eventos if e["id"]==eid)
                    with st.form(f"editar_{eid}"):
                        en=st.text_input("Nome",evt["nome"]);el=st.text_input("Local",evt["local"]);ed=st.selectbox("Dia",range(7),index=int(evt["dia_sem"]),format_func=lambda x:DIAS_SEMANA_PT[x]);es=st.selectbox("Semana",[1,2,3,4,5],index=int(evt["semana"])-1);eh=st.text_input("Hora",evt["hora"]);ef=st.selectbox("Frequência",FREQUENCIAS,index=FREQUENCIAS.index(evt["interc"]))
                        if st.form_submit_button("✅ SALVAR ALTERAÇÕES",use_container_width=True):
                            if not el.strip():st.error("Informe o local.")
                            else:
                                try:atualizar_evento(eid,{"nome":en.strip().upper(),"local":el.strip().upper(),"dia_sem":ed,"semana":es,"hora":eh.strip().upper(),"interc":ef});definir_notificacao("✅ Alterações salvas com sucesso!");st.rerun()
                                except Exception as e:st.error(f"Erro ao atualizar: {e}")
                    st.divider();st.warning(f"Excluir permanentemente: {evt['nome']} — {evt['local']}");conf=st.checkbox("Confirmo que desejo excluir este evento",key=f"confirmar_exclusao_{eid}")
                    if st.button("🗑️ EXCLUIR EVENTO",use_container_width=True,disabled=not conf,key=f"excluir_{eid}"):
                        try:excluir_evento(eid);definir_notificacao("✅ Evento excluído com sucesso!");st.rerun()
                        except Exception as e:st.error(f"Erro ao excluir: {e}")
            with abas[3]:
                try:st.download_button("⬇️ Baixar Excel",gerar_excel(st.session_state.ano_base,eventos,avisos),f"Calendario_{st.session_state.ano_base}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                except Exception as e:st.error(f"Não foi possível gerar o Excel: {e}")
                try:st.download_button("⬇️ Baixar PDF",gerar_pdf(st.session_state.ano_base,eventos,avisos),f"Calendario_{st.session_state.ano_base}.pdf",mime="application/pdf",use_container_width=True)
                except Exception as e:st.error(f"Não foi possível gerar o PDF: {e}")
        elif senha:st.error("❌ Senha incorreta")
    st.markdown('</div>',unsafe_allow_html=True)
