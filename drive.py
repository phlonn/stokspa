"""Accès à Google Drive : les données du dépôt sont rangées dans le dossier caché propre à
l'application (« appDataFolder »). Elles sont ainsi les mêmes sur le PC et dans la version en
ligne, et ne peuvent pas être supprimées par erreur depuis Drive."""

from __future__ import annotations

import io
import json
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

BASE_DIR = Path(__file__).parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"  # client OAuth "Application de bureau"
TOKEN_FILE = BASE_DIR / "token.json"  # créé automatiquement après la 1re connexion
SCOPES = ["https://www.googleapis.com/auth/drive.appdata"]  # dossier caché de l'appli uniquement


class DriveError(RuntimeError):
    pass


def secret(name: str):
    """Valeur des secrets Streamlit (version en ligne), ou None sur le PC."""
    try:
        import streamlit as st

        return st.secrets.get(name)
    except Exception:
        return None


def _token_secret() -> dict | None:
    t = secret("google_token")
    if not t:
        return None
    return json.loads(t) if isinstance(t, str) else dict(t)


def en_ligne() -> bool:
    """Vrai pour la version hébergée (le jeton Google vient des secrets)."""
    return _token_secret() is not None


def _credentials() -> Credentials:
    info = _token_secret()
    creds = None
    if info:
        creds = Credentials.from_authorized_user_info(info)
    elif TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE))
    if creds and not creds.has_scopes(SCOPES):
        creds = None
    if creds and creds.valid:
        return creds
    if creds and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None
    if creds and creds.valid:
        if not info:
            TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
        return creds
    if info:
        raise DriveError(
            "L'accès à Google Drive a expiré. Lancez lancer.bat sur le PC, autorisez Google Drive, "
            "puis relancez preparer_en_ligne.py et recopiez les secrets dans Streamlit (voir README)."
        )
    if not CREDENTIALS_FILE.exists():
        raise DriveError("Fichier credentials.json introuvable (voir README).")
    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
    creds = flow.run_local_server(port=0, prompt="consent")  # ouvre le navigateur
    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return creds


def _service():
    return build("drive", "v3", credentials=_credentials(), cache_discovery=False)


def _appdata_id(svc, name: str) -> str | None:
    res = svc.files().list(spaces="appDataFolder", q=f"name = '{name}' and trashed = false", fields="files(id)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def read_json(name: str) -> dict | None:
    """Contenu du fichier JSON du dossier caché, ou None s'il n'existe pas encore."""
    svc = _service()
    fid = _appdata_id(svc, name)
    if not fid:
        return None
    return json.loads(svc.files().get_media(fileId=fid).execute())


def write_json(name: str, data: dict) -> None:
    svc = _service()
    media = MediaIoBaseUpload(io.BytesIO(json.dumps(data, ensure_ascii=False).encode("utf-8")), mimetype="application/json")
    fid = _appdata_id(svc, name)
    if fid:
        svc.files().update(fileId=fid, media_body=media).execute()
    else:
        svc.files().create(body={"name": name, "parents": ["appDataFolder"]}, media_body=media, fields="id").execute()
