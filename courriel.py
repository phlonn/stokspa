"""Envoi d'un e-mail avec pièce jointe, par le serveur SMTP de Gmail (mot de passe d'application)."""

from __future__ import annotations

import base64
import re
import smtplib
import ssl
from email.message import EmailMessage

ADRESSE = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")


class CourrielError(RuntimeError):
    pass


def destinataires(texte: str) -> list[str]:
    adresses = [a.strip() for a in re.split(r"[,;\s]+", texte or "") if a.strip()]
    if not adresses:
        raise CourrielError("Indiquez au moins une adresse e-mail.")
    fausses = [a for a in adresses if not ADRESSE.match(a)]
    if fausses:
        raise CourrielError("Adresse invalide : " + ", ".join(fausses))
    if len(adresses) > 20:
        raise CourrielError("20 destinataires au maximum.")
    return adresses


def envoyer(cfg: dict, a: str, objet: str, texte: str, fichier: dict) -> list[str]:
    """cfg : utilisateur, mot_de_passe, hote, port. fichier : {"name", "data" (base64)}. Renvoie les destinataires."""
    if not cfg.get("utilisateur") or not cfg.get("mot_de_passe"):
        raise CourrielError("L'envoi d'e-mails n'est pas configuré : voir « Envoi du stock par e-mail » dans le README.")
    dest = destinataires(a)
    m = EmailMessage()
    m["From"] = cfg["utilisateur"]
    m["To"] = ", ".join(dest)
    m["Subject"] = (objet or "Stock du dépôt").strip()[:200]
    m.set_content(texte or "")
    try:
        contenu = base64.b64decode(fichier["data"], validate=True)
    except Exception as e:
        raise CourrielError("Pièce jointe illisible.") from e
    m.add_attachment(
        contenu,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=re.sub(r"[^\w.\-]", "_", fichier.get("name") or "stock.xlsx"),
    )
    hote, port = cfg.get("hote") or "smtp.gmail.com", int(cfg.get("port") or 465)
    try:
        if port == 465:
            with smtplib.SMTP_SSL(hote, port, context=ssl.create_default_context(), timeout=30) as s:
                s.login(cfg["utilisateur"], cfg["mot_de_passe"])
                s.send_message(m)
        else:
            with smtplib.SMTP(hote, port, timeout=30) as s:
                if s.has_extn("starttls"):
                    s.starttls(context=ssl.create_default_context())
                if s.has_extn("auth"):
                    s.login(cfg["utilisateur"], cfg["mot_de_passe"])
                s.send_message(m)
    except smtplib.SMTPAuthenticationError as e:
        raise CourrielError("Gmail a refusé la connexion : vérifiez l'adresse et le mot de passe d'application.") from e
    except (smtplib.SMTPException, OSError) as e:
        raise CourrielError(f"L'e-mail n'a pas pu être envoyé : {e}") from e
    return dest
