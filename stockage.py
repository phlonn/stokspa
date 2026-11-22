"""Données du dépôt (produits, emplacements, stock, mouvements, réglages), enregistrées dans
Google Drive. Une copie est gardée sur le PC (data/stock.json) quand l'application tourne en local."""

from __future__ import annotations

import json
from pathlib import Path

import drive

FICHIER = "donnees_depot.json"
LOCAL = Path(__file__).parent / "data" / "stock.json"
COLLECTIONS = ("products", "locations", "stock", "movements")


def _vide() -> dict:
    return {**{c: {} for c in COLLECTIONS}, "settings": {"alertMonths": 2}}


def _normaliser(d: dict) -> dict:
    base = _vide()
    for c in COLLECTIONS:
        base[c] = dict(d.get(c) or {})
    base["settings"] = {**base["settings"], **(d.get("settings") or {})}
    return base


def charger() -> dict:
    """Lit les données dans Google Drive. La première fois, part de la copie du PC."""
    try:
        d = drive.read_json(FICHIER)
    except drive.DriveError:
        raise
    except Exception as e:
        raise drive.DriveError(f"Google Drive injoignable : {e}") from e
    if d is None:
        d = json.loads(LOCAL.read_text(encoding="utf-8")) if LOCAL.exists() else _vide()
        d = _normaliser(d)
        enregistrer(d)
        return d
    return _normaliser(d)


def enregistrer(d: dict) -> None:
    drive.write_json(FICHIER, d)
    if not drive.en_ligne():  # copie de secours sur le PC
        LOCAL.parent.mkdir(exist_ok=True)
        LOCAL.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def appliquer(ops: list[dict]) -> dict:
    """Relit Drive (pour garder les changements faits ailleurs), applique les opérations, enregistre."""
    d = charger()
    for o in ops:
        if "settings" in o:
            d["settings"].update(o["settings"] or {})
            continue
        c, i = o.get("c"), str(o.get("id") or "")
        if c not in COLLECTIONS or not i:
            continue
        if o.get("del"):
            d[c].pop(i, None)
        else:
            d[c][i] = {**(o.get("data") or {}), "id": i}
    enregistrer(d)
    return d
