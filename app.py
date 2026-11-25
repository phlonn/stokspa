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

import courriel
import drive
import imprimante
import lecteur
import scanner
import stockage

BASE = Path(__file__).parent
st.set_page_config(page_title="FabLog", page_icon="💧", layout="wide")

# --- Mot de passe (version en ligne) ---------------------------------------------
MDP = drive.secret("mot_de_passe")
if MDP and not st.session_state.get("autorise"):
    st.title("💧 FabLog")
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


def reglage(nom: str) -> str:
    """Valeur des secrets Streamlit (en ligne), ou de secrets_streamlit.toml sur le PC."""
    m = drive.secret(nom)
    if m:
        return str(m)
    f = BASE / "secrets_streamlit.toml"
    if f.exists():
        import tomllib

        try:
            v = tomllib.loads(f.read_text(encoding="utf-8-sig")).get(nom, "")
            return "" if str(v).startswith("CHOISISSEZ") else str(v)
        except Exception:
            return ""
    return ""


def mdp_admin() -> str:
    """Mot de passe de « Gestion des données »."""
    return reglage("mot_de_passe_admin")


def config_mail() -> dict:
    return {
        "utilisateur": reglage("smtp_utilisateur"),
        "mot_de_passe": reglage("smtp_mot_de_passe").replace(" ", ""),  # Google l'affiche par groupes de 4
        "hote": reglage("smtp_hote"),
        "port": reglage("smtp_port"),
    }


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
    ss.rev, ss.ack, ss.ocr, ss.erreur, ss.dernier, ss.admin, ss.mail = 1, None, None, None, None, None, None

_cm = config_mail()
etat_mail = {"configured": bool(_cm["utilisateur"] and _cm["mot_de_passe"]), "from": _cm["utilisateur"], "res": ss.get("mail")}
depot = components.declare_component("depot", path=str(BASE / "composant"))
msg = depot(data=ss.donnees, rev=ss.rev, ack=ss.get("ack"), ocr=ss.get("ocr"), error=ss.get("erreur"), admin=ss.get("admin"), mail=etat_mail,
            scan={"available": scanner.DISPONIBLE, "res": ss.get("scan")},
            imp={"available": imprimante.DISPONIBLE, "res": ss.get("imp")}, key="depot", default=None)

# --- Messages envoyés par la page ---------------------------------------------------
if isinstance(msg, dict) and msg.get("id") and msg["id"] != ss.dernier:
    ss.dernier = msg["id"]
    genre = msg.get("kind")
    ss.scan = None  # l'image numérisée n'est envoyée qu'une fois à la page
    try:
        if genre == "ops":
            ss.donnees = stockage.appliquer(msg.get("ops") or [])
            ss.rev += 1
        elif genre == "sync":
            nouveau = stockage.charger()
            if nouveau != ss.donnees:
                ss.donnees = nouveau
                ss.rev += 1
        elif genre == "admin":
            attendu = mdp_admin()
            if not attendu or attendu.startswith("CHOISISSEZ"):
                ss.admin = {"id": msg["id"], "ok": False, "msg": "Aucun mot de passe n'est défini pour la gestion des données (voir README)."}
            else:
                ok = hmac.compare_digest(str(msg.get("pwd", "")).encode(), attendu.encode())
                ss.admin = {"id": msg["id"], "ok": ok, "msg": "" if ok else "Mot de passe incorrect."}
        elif genre == "mail":
            try:
                dest = courriel.envoyer(config_mail(), msg.get("to", ""), msg.get("subject", ""), msg.get("text", ""), msg.get("file") or {})
                ss.mail = {"id": msg["id"], "ok": True, "msg": "E-mail envoyé à " + ", ".join(dest) + "."}
            except courriel.CourrielError as e:
                ss.mail = {"id": msg["id"], "ok": False, "msg": str(e)}
        elif genre == "scan":
            try:
                img = scanner.numeriser()
                ss.scan = {"id": msg["id"], "ok": True, "image": img} if img else {"id": msg["id"], "ok": False, "cancelled": True}
            except scanner.ScanError as e:
                ss.scan = {"id": msg["id"], "ok": False, "msg": str(e)}
        elif genre == "print":
            try:
                nom = imprimante.imprimer_xlsx((msg.get("file") or {}).get("data", ""), msg.get("title", ""), msg.get("infos") or [])
                ss.imp = {"id": msg["id"], "ok": True, "msg": f"Liste envoyée à l'imprimante {nom}." if nom else "Liste envoyée à l'imprimante."}
            except imprimante.ImpressionError as e:
                ss.imp = {"id": msg["id"], "ok": False, "msg": str(e)}
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
