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

def solicitar_acesso_inicial():
    email = "lucasduarteccb123@hotmail.com"
    st.info("Primeiro acesso: defina sua senha de administrador.")
    with st.form("primeiro_acesso"):
        senha = st.text_input("Criar senha", type="password")
        confirmar = st.text_input("Confirmirmar senha", type="password")
        criar = st.form_submit_button("ATIVAR MEU ACESSO", use_container_width=True, type="primary")
    if criar:
        if len(senha) < 8:
            st.error("Use uma senha com pelo menos 8 caracteres.")
        elif senha != confirmar:
            st.error("As senhas não conferem.")
        else:
            r = requests.post(f"{SUPABASE_URL}/auth/v1/signup", headers=_headers(), json={"email": email, "password": senha, "data": {"nome": "Lucas Duarte"}}, timeout=15)
            if r.ok:
                st.success("Conta criada. Confira seu e-mail para confirmar o acesso e depois entre normalmente.")
            else:
                try:
                    msg = r.json().get("msg") or r.json().get("message")
                except Exception:
                    msg = None
                st.error(msg or "Não foi possível ativar o acesso.")

def exigir_login():
    if st.session_state.get("access_token"):
        return True
    st.markdown("## 🔐 Agenda Musical")
    st.caption("Acesso restrito • Região de Jaciara - MT")
    with st.form("login_global"):
        email = st.text_input("Usuário / E-mail")
        senha = st.text_input("Senha", type="password")
        entrar = st.form_submit_button("ENTRAR", use_container_width=True, type="primary")
    if entrar:
        ok, erro = login(email, senha)
        if ok:
            st.rerun()
        st.error(erro)
    st.divider()
    solicitar_acesso_inicial()
    st.stop()

def eh_admin():
    return (st.session_state.get("perfil") or {}).get("papel") == "admin"

def token():
    return st.session_state.get("access_token", "")
