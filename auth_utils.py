import requests
import streamlit as st

SUPABASE_URL = "https://ovnwnzqjjjtfqjodvusi.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_uBqke5HDz9U-xSKjxhzUww_-Y0qW367"

def _headers(token=None):
    key = SUPABASE_PUBLISHABLE_KEY
    return {"apikey": key, "Authorization": f"Bearer {token or key}", "Content-Type": "application/json"}

def carregar_perfil():
    access_token = st.session_state.get("access_token")
    user = st.session_state.get("auth_user") or {}
    uid = user.get("id")
    if not access_token or not uid:
        return None
    r = requests.get(f"{SUPABASE_URL}/rest/v1/perfis", headers=_headers(access_token), params={"user_id": f"eq.{uid}", "select": "user_id,nome,papel,ativo"}, timeout=10)
    data = r.json() if r.ok else []
    perfil = data[0] if data else {"user_id": uid, "nome": user.get("email", ""), "papel": "usuario", "ativo": True}
    st.session_state["perfil"] = perfil
    return perfil

def login(email, senha):
    r = requests.post(f"{SUPABASE_URL}/auth/v1/token?grant_type=password", headers=_headers(), json={"email": email.strip(), "password": senha}, timeout=15)
    if not r.ok:
        return False, "Usuário ou senha inválidos."
    data = r.json()
    st.session_state["access_token"] = data["access_token"]
    st.session_state["refresh_token"] = data.get("refresh_token")
    st.session_state["auth_user"] = data.get("user", {})
    carregar_perfil()
    return True, None

def logout():
    access_token = st.session_state.get("access_token")
    if access_token:
        try:
            requests.post(f"{SUPABASE_URL}/auth/v1/logout", headers=_headers(access_token), timeout=10)
        except Exception:
            pass
    for key in ("access_token", "refresh_token", "auth_user", "perfil"):
        st.session_state.pop(key, None)

def exigir_login():
    if st.session_state.get("access_token"):
        return True
    st.markdown("""
    <style>
    [data-testid="stSidebar"]{display:none}
    .block-container{max-width:560px;padding-top:7vh}
    .login-brand{text-align:center;padding:28px 10px 18px}
    .login-icon{font-size:48px}
    .login-title{font-family:Georgia,serif;font-size:38px;font-weight:800;color:#082d49;margin:8px 0 2px}
    .login-sub{color:#b17b00;font-weight:700}
    .login-info{text-align:center;color:#64748b;margin:20px 0}
    </style>
    <div class="login-brand">
      <div class="login-icon">🎼</div>
      <div class="login-title">Agenda Musical</div>
      <div class="login-sub">Região de Jaciara - MT</div>
    </div>
    <div class="login-info">Acesse a agenda, registros de ensaios e aulas do MSA em um só lugar.</div>
    """, unsafe_allow_html=True)
    with st.form("login_global"):
        email=st.text_input("E-mail",placeholder="seuemail@exemplo.com")
        senha=st.text_input("Senha",type="password",placeholder="Digite sua senha")
        entrar=st.form_submit_button("ENTRAR",use_container_width=True,type="primary")
    if entrar:
        ok,erro=login(email,senha)
        if ok:
            st.rerun()
        st.error(erro)
    st.caption("🔒 Acesso restrito a usuários autorizados.")
    st.stop()

def eh_admin():
    return (st.session_state.get("perfil") or {}).get("papel") == "admin"

def token():
    return st.session_state.get("access_token", "")
