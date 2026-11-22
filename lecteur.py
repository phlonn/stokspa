"""Lecture par Claude des documents importés (relevé de stock manuscrit, bon de commande)."""

from __future__ import annotations

import json
import re

import anthropic

MODEL = "claude-opus-5-5"


class LectureError(RuntimeError):
    pass


def lire_json(prompt: str, images: list[dict], api_key: str):
    """images : [{"media_type": "image/jpeg", "data": "<base64>"}]. Renvoie le JSON lu par Claude."""
    if not api_key:
        raise LectureError("Clé API Anthropic manquante : ajoutez-la dans config.json (PC) ou dans les secrets Streamlit.")
    contenu = [
        {"type": "image", "source": {"type": "base64", "media_type": i.get("media_type", "image/jpeg"), "data": i["data"]}}
        for i in images[:20]
    ]
    contenu.append({"type": "text", "text": prompt})
    try:
        r = anthropic.Anthropic(api_key=api_key).messages.create(
            model=MODEL, max_tokens=16000, messages=[{"role": "user", "content": contenu}]
        )
    except anthropic.RateLimitError as e:
        raise LectureError("Trop de demandes pour le moment, réessayez dans quelques minutes.") from e
    except anthropic.AuthenticationError as e:
        raise LectureError("Clé API Anthropic refusée : vérifiez-la.") from e
    except anthropic.APIError as e:
        raise LectureError(f"La lecture a échoué : {e}") from e
    if r.stop_reason == "refusal":
        raise LectureError("Claude n'a pas pu lire ce document.")
    texte = "".join(b.text for b in r.content if b.type == "text").strip()
    texte = re.sub(r"^```(?:json)?\s*|\s*```$", "", texte)
    try:
        return json.loads(texte)
    except json.JSONDecodeError:
        m = re.search(r"(\[.*\]|\{.*\})", texte, re.S)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass
    raise LectureError("La réponse n'a pas pu être lue, réessayez.")
