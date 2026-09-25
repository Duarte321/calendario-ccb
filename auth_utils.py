import base64
import hashlib
import hmac
import secrets
import time
from urllib.parse import quote

import requests
import streamlit as st

SUPABASE_URL = "https://ovnwnzqjjjtfqjodvusi.supabase.co"
SUPABASE_PUBLISHABLE_KEY = "sb_publishable_uBqke5HDz9U-xSKjxhzUww_-Y0qW367"
APP_URL = "https://calendario-ccb-4dnxiyyyxfshae2pfswcwf.streamlit.app"

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

def _b64url(data):
    return base64.urlsafe_b64encode(data).decode("utf-8").rstrip("=")

def _recovery_secret():
    return st.secrets.get("RECOVERY_SECRET", "") or st.secrets.get("SUPABASE_SECRET_KEY", "")

def _recovery_verifier(recovery_id):
    segredo=_recovery_secret()
    if not segredo:
        return None
    digest=hmac.new(segredo.encode("utf-8"),recovery_id.encode("utf-8"),hashlib.sha256).digest()
    return _b64url(digest)

def enviar_recuperacao(email):
    email=(email or "").strip()
    if not email:
        return False,"Informe seu e-mail."
    verifier_id=secrets.token_urlsafe(18)
    verifier=_recovery_verifier(verifier_id)
    if not verifier:
        return False,"Recuperação de senha ainda não está configurada no servidor."
    challenge=_b64url(hashlib.sha256(verifier.encode("utf-8")).digest())
    redirect_to=f"{APP_URL}/?recovery_id={quote(verifier_id)}"
    r=requests.post(
        f"{SUPABASE_URL}/auth/v1/recover",
        params={"redirect_to":redirect_to},
        headers=_headers(),
        json={"email":email,"code_challenge":challenge,"code_challenge_method":"s256"},
        timeout=15,
    )
    if r.ok:
        return True,"Se esse e-mail estiver cadastrado, você receberá um link para redefinir sua senha."
    if r.status_code==429:
        return False,"Aguarde um pouco antes de solicitar outro e-mail de recuperação."
    return False,"Não foi possível enviar o e-mail de recuperação agora."

def _processar_callback_recuperacao():
    try:
        code=st.query_params.get("code")
        recovery_id=st.query_params.get("recovery_id")
    except Exception:
        return
    if not code or not recovery_id:
        return
    verifier=_recovery_verifier(recovery_id)
    if not verifier:
        st.session_state["recovery_error"]="Não foi possível validar a recuperação de senha."
        return
    r=requests.post(
        f"{SUPABASE_URL}/auth/v1/token?grant_type=pkce",
        headers=_headers(),
        json={"auth_code":code,"code_verifier":verifier},
        timeout=15,
    )
    try:
        st.query_params.clear()
    except Exception:
        pass
    if r.ok:
        data=r.json()
        st.session_state["recovery_access_token"]=data.get("access_token")
        st.session_state["recovery_error"]=None
    else:
        st.session_state["recovery_error"]="O link de recuperação é inválido ou expirou."

def _tela_nova_senha():
    access_token=st.session_state.get("recovery_access_token")
    if not access_token:
        return False
    st.markdown("## 🔐 Criar nova senha")
    st.caption("Digite uma nova senha para sua conta.")
    with st.form("nova_senha_form"):
        nova=st.text_input("Nova senha",type="password")
        confirmar=st.text_input("Confirmar nova senha",type="password")
        salvar=st.form_submit_button("SALVAR NOVA SENHA",use_container_width=True,type="primary")
    if salvar:
        if len(nova)<8:
            st.error("Use uma senha com pelo menos 8 caracteres.")
        elif nova!=confirmar:
            st.error("As senhas não conferem.")
        else:
            r=requests.put(
                f"{SUPABASE_URL}/auth/v1/user",
                headers={"apikey":SUPABASE_PUBLISHABLE_KEY,"Authorization":f"Bearer {access_token}","Content-Type":"application/json"},
                json={"password":nova},
                timeout=15,
            )
            if r.ok:
                st.session_state.pop("recovery_access_token",None)
                st.session_state["password_reset_success"]=True
                st.rerun()
            else:
                st.error("Não foi possível alterar a senha. Solicite um novo link de recuperação.")
    st.stop()

def exigir_login():
    _processar_callback_recuperacao()
    _tela_nova_senha()

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

    if st.session_state.pop("password_reset_success",False):
        st.success("Senha alterada com sucesso. Entre com sua nova senha.")

    erro_rec=st.session_state.pop("recovery_error",None)
    if erro_rec:
        st.error(erro_rec)

    modo=st.session_state.get("login_mode","login")

    if modo=="recovery":
        st.markdown("### Esqueci minha senha")
        st.caption("Informe o e-mail da sua conta. Enviaremos um link seguro para você criar uma nova senha.")

        agora=time.time()
        ultimo_envio=st.session_state.get("recovery_last_sent_at",0)
        restante=max(0,int(60-(agora-ultimo_envio))) if ultimo_envio else 0

        with st.form("recovery_form"):
            email_rec=st.text_input("E-mail",placeholder="seuemail@exemplo.com")
            enviar=st.form_submit_button(
                "ENVIAR LINK DE RECUPERAÇÃO" if restante==0 else f"AGUARDE {restante}s",
                use_container_width=True,
                type="primary",
                disabled=restante>0,
            )

        if enviar and restante==0:
            ok,msg=enviar_recuperacao(email_rec)
            if ok:
                st.session_state["recovery_last_sent_at"]=time.time()
                st.session_state["recovery_notice"]=msg
                st.session_state.pop("recovery_rate_error",None)
                st.rerun()
            else:
                if "Aguarde" in msg:
                    st.session_state["recovery_last_sent_at"]=time.time()
                    st.session_state["recovery_rate_error"]="O Supabase limitou temporariamente novos e-mails. Aguarde 60 segundos e tente novamente."
                    st.rerun()
                st.error(msg)

        aviso=st.session_state.get("recovery_notice")
        if aviso:
            st.success(aviso)

        erro_limite=st.session_state.get("recovery_rate_error")
        if erro_limite:
            st.warning(erro_limite)

        agora=time.time()
        ultimo_envio=st.session_state.get("recovery_last_sent_at",0)
        restante=max(0,int(60-(agora-ultimo_envio))) if ultimo_envio else 0
        if restante>0:
            st.info(f"Você poderá solicitar outro link em aproximadamente {restante} segundos.")
            st.caption("Recarregue a página para atualizar a contagem.")
        elif erro_limite:
            st.session_state.pop("recovery_rate_error",None)

        if st.button("← Voltar para o login",use_container_width=True):
            st.session_state["login_mode"]="login"
            st.rerun()
        st.stop()

    with st.form("login_global"):
        email=st.text_input("E-mail",placeholder="seuemail@exemplo.com")
        senha=st.text_input("Senha",type="password",placeholder="Digite sua senha")
        entrar=st.form_submit_button("ENTRAR",use_container_width=True,type="primary")
    if entrar:
        ok,erro=login(email,senha)
        if ok:
            st.rerun()
        st.error(erro)

    if st.button("🔑 Esqueci minha senha",use_container_width=True):
        st.session_state["login_mode"]="recovery"
        st.rerun()

    st.caption("🔒 Acesso restrito a usuários autorizados.")
    st.stop()

def eh_admin():
    return (st.session_state.get("perfil") or {}).get("papel") == "admin"

def token():
    return st.session_state.get("access_token", "")
