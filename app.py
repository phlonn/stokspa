"""Gestion de dépôt – palettes d'eau.

Lancement sur le PC :  lancer.bat   —   version en ligne : Streamlit Community Cloud (voir README).
L'interface est la page « composant/index.html » ; ce script la sert, la protège par mot de passe
et enregistre les données dans Google Drive.
"""

from __future__ import annotations

import hmac
import json
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import drive
import lecteur
import stockage

BASE = Path(__file__).parent
st.set_page_config(page_title="Dépôt – Palettes d'eau", page_icon="💧", layout="wide")

# --- Mot de passe (version en ligne) ---------------------------------------------
MDP = drive.secret("mot_de_passe")
if MDP and not st.session_state.get("autorise"):
    st.title("💧 Dépôt – Palettes d'eau")
    saisie = st.text_input("Mot de passe", type="password")
    if saisie:
        if hmac.compare_digest(saisie.encode(), str(MDP).encode()):
            st.session_state.autorise = True
            st.rerun()
        st.error("Mot de passe incorrect.")
    st.stop()


def cle_api() -> str:
    cle = drive.secret("anthropic_api_key")
    if cle:
        return str(cle)
    cfg = BASE / "config.json"
    if cfg.exists():
        return json.loads(cfg.read_text(encoding="utf-8-sig")).get("anthropic_api_key", "")
    return ""


# --- Pleine page : la page du dépôt occupe toute la fenêtre ------------------------
st.markdown(
    """<style>
    header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], footer {display:none !important}
    .block-container, [data-testid="stMainBlockContainer"] {padding:0 !important; max-width:100% !important}
    [data-testid="stAppViewContainer"] {overflow:hidden}
    iframe {position:fixed !important; inset:0; width:100vw !important; height:100vh !important; height:100dvh !important; border:0; z-index:1000}
    </style>""",
    unsafe_allow_html=True,
)

ss = st.session_state
if "donnees" not in ss:
    try:
        with st.spinner("Chargement du dépôt depuis Google Drive…"):
            ss.donnees = stockage.charger()
    except drive.DriveError as e:
        st.error(str(e))
        st.stop()
    ss.rev, ss.ack, ss.ocr, ss.erreur, ss.dernier = 1, None, None, None, None

depot = components.declare_component("depot", path=str(BASE / "composant"))
msg = depot(data=ss.donnees, rev=ss.rev, ack=ss.ack, ocr=ss.ocr, error=ss.erreur, key="depot", default=None)

# --- Messages envoyés par la page ---------------------------------------------------
if isinstance(msg, dict) and msg.get("id") and msg["id"] != ss.dernier:
    ss.dernier = msg["id"]
    genre = msg.get("kind")
    try:
        if genre == "ops":
            ss.donnees = stockage.appliquer(msg.get("ops") or [])
            ss.rev += 1
        elif genre == "sync":
            nouveau = stockage.charger()
            if nouveau != ss.donnees:
                ss.donnees = nouveau
                ss.rev += 1
        elif genre == "ocr":
            try:
                ss.ocr = {"id": msg["id"], "result": lecteur.lire_json(msg.get("prompt", ""), msg.get("images") or [], cle_api())}
            except lecteur.LectureError as e:
                ss.ocr = {"id": msg["id"], "error": {"code": "server", "msg": str(e)}}
    except Exception as e:  # Drive injoignable : la page revient aux dernières données enregistrées
        ss.erreur = {"id": msg["id"], "msg": f"Enregistrement dans Google Drive impossible : {e}"}
        try:
            ss.donnees = stockage.charger()
        except Exception:
            pass
        ss.rev += 1
    ss.ack = msg["id"]
    st.rerun()
