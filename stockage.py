"""Base de données des offres : un fichier JSON par offre + journal des événements (CSV).

data/
  offres/2c/2c5f…json   une offre : métadonnées brutes (liste + fiche) et horodatages de suivi
  historique.csv        mises en ligne, retraits, remises en ligne, modifications
  etat.json             état général du suivi

Un fichier par offre garde l'historique Git lisible : seules les offres qui changent sont réécrites.
"""

from __future__ import annotations

import csv
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from annonce import PARIS, dig, fmt, labels, pick

EVENT_COLUMNS = ["horodatage", "evenement", "guid", "reference", "ville", "code_postal",
                 "typologie", "loyer", "details"]

# Champs dont un changement est noté dans l'historique (chemin, libellé, format)
TRACKED = [
    ("totalRentAmountBaseBound", "Loyer", "money"),
    ("totalProgressiveRentAmount", "Loyer progressif", "money"),
    ("estimatedExternalCharges", "Charges estimées", "money"),
    ("agencyFeesAmount", "Frais d'agence", "money"),
    ("availableAt", "Disponible le", "date"),
    ("startDateOfPublication", "Publication", "date"),
    ("lodging.area", "Surface", "area"),
    ("lodging.lodgingTypology.label", "Typologie", "text"),
]
VOLATILE = {"isRecent"}  # change tout seul avec le temps : ignoré pour la détection de modifications


def now_iso() -> str:
    return datetime.now(PARIS).isoformat(timespec="seconds")


def clean(obj: Any) -> Any:
    """Retire la partie temporaire (signature, expiration) des liens de photos."""
    if isinstance(obj, dict):
        out = {}
        for k, v in obj.items():
            if k == "url" and isinstance(v, str) and "X-Amz-" in v:
                v = v.split("?", 1)[0]
            out[k] = clean(v)
        return out
    if isinstance(obj, list):
        return [clean(x) for x in obj]
    return obj


# Adresses mail nominatives (prenom.nom@…) parfois écrites par les bailleurs dans leurs annonces :
# masquées avant tout enregistrement, puisque le dépôt et les Excel sont publics.
PERSON_EMAIL = re.compile(r"\b[A-Za-zÀ-ÿ]+(?:[.\-_][A-Za-zÀ-ÿ]+)+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")
MASK = "[adresse masquée]"


def mask_people(obj: Any) -> Any:
    """Masque en place les adresses nominatives dans toutes les chaînes d'une offre."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, str):
                obj[k] = PERSON_EMAIL.sub(MASK, v)
            else:
                mask_people(v)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            if isinstance(v, str):
                obj[i] = PERSON_EMAIL.sub(MASK, v)
            else:
                mask_people(v)
    return obj


def comparable(offer: dict) -> dict:
    return {k: v for k, v in offer.items() if k not in VOLATILE}


def al_summary(rec: dict) -> dict:
    """Colonnes de l'historique pour une offre Action Logement."""
    return {
        "reference": pick(rec, "reference") or "",
        "ville": pick(rec, "lodging.municipality.name") or "",
        "code_postal": pick(rec, "lodging.municipality.postcode") or "",
        "typologie": labels(pick(rec, "lodging.lodgingTypology")) or "",
        "loyer": fmt(pick(rec, "totalRentAmountBaseBound"), "money") or "",
    }


class Store:
    def __init__(self, data_dir: Path, summary=None, tracked=None):
        self.summary = summary or al_summary   # colonnes de l'historique, selon le site
        self.tracked = tracked if tracked is not None else TRACKED
        self.dir = Path(data_dir)
        self.offers_dir = self.dir / "offres"
        self.history_path = self.dir / "historique.csv"
        self.state_path = self.dir / "etat.json"
        self.records: dict[str, dict] = {}
        self._dirty: set[str] = set()
        self._events: list[dict] = []
        self.state = {"initialise": False, "debut_suivi": None, "dernier_passage": None,
                      "dernier_passage_complet": None, "passages_vides": 0, "alerte_vide_envoyee": False}
        self._state_dirty = False
        self._load()

    # -- chargement / sauvegarde ------------------------------------------------------
    def _load(self) -> None:
        if self.state_path.exists():
            self.state.update(json.loads(self.state_path.read_text(encoding="utf-8")))
        if self.offers_dir.exists():
            for path in self.offers_dir.glob("*/*.json"):
                rec = json.loads(path.read_text(encoding="utf-8"))
                self.records[rec["guid"]] = rec

    def _path(self, guid: str) -> Path:
        safe = re.sub(r"[^0-9A-Za-z-]", "_", guid)
        return self.offers_dir / safe[:2] / f"{safe}.json"

    def touch(self, guid: str) -> None:
        self._dirty.add(guid)

    def set_state(self, **kwargs) -> None:
        for k, v in kwargs.items():
            if self.state.get(k) != v:
                self.state[k] = v
                self._state_dirty = True

    @property
    def changed(self) -> bool:
        return bool(self._dirty or self._events)

    def save(self) -> bool:
        """Écrit ce qui a changé ; renvoie True si les données des offres ont changé."""
        data_changed = bool(self._dirty or self._events)
        for guid in sorted(self._dirty):
            mask_people(self.records[guid])
            path = self._path(guid)
            path.parent.mkdir(parents=True, exist_ok=True)
            tmp = path.with_suffix(".tmp")
            tmp.write_text(json.dumps(self.records[guid], ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
            tmp.replace(path)
        if self._events:
            new_file = not self.history_path.exists()
            self.history_path.parent.mkdir(parents=True, exist_ok=True)
            with self.history_path.open("a", encoding="utf-8-sig" if new_file else "utf-8", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=EVENT_COLUMNS, delimiter=";")
                if new_file:
                    writer.writeheader()
                writer.writerows(self._events)
        if self._state_dirty or self._dirty or self._events:
            self.dir.mkdir(parents=True, exist_ok=True)
            self.state_path.write_text(json.dumps(self.state, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
        self._dirty.clear()
        self._events.clear()
        self._state_dirty = False
        return data_changed

    def read_history(self) -> list[dict]:
        if not self.history_path.exists():
            return []
        with self.history_path.open(encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f, delimiter=";"))

    # -- événements -------------------------------------------------------------------
    def event(self, when: str, kind: str, rec: dict | None, details: str = "") -> None:
        info = self.summary(rec) if rec else {}
        self._events.append({
            "horodatage": when, "evenement": kind, "guid": (rec or {}).get("guid", ""),
            **{k: info.get(k, "") for k in ("reference", "ville", "code_postal", "typologie", "loyer")},
            "details": details,
        })

    # -- mise à jour à partir d'un passage --------------------------------------------
    def new_record(self, offer: dict, when: str, initial: bool, in_zone: bool) -> dict:
        rec = {
            "guid": offer["guid"],
            "statut": "en_ligne",
            "premiere_detection": when,
            "deja_en_ligne_au_demarrage": initial,
            "derniere_presence": None,
            "absente_depuis": None,
            "absences": 0,
            "retrait_detecte": None,
            "remises_en_ligne": 0,
            "alerte": "a_envoyer" if in_zone and not initial else None,
            "alerte_envoyee_le": None,
            "fiche_statut": None,
            "fiche_recuperee_le": None,
            "liste": clean(offer),
            "fiche": None,
        }
        self.records[rec["guid"]] = rec
        self.touch(rec["guid"])
        if not initial:
            self.event(when, "Mise en ligne", rec)
        return rec

    def update_listing(self, rec: dict, offer: dict, when: str) -> list[str]:
        """Met à jour la version « liste » si elle a changé ; renvoie les modifications notables."""
        offer = mask_people(clean(offer))   # comparée à la version enregistrée, déjà masquée
        old = rec.get("liste") or {}
        if comparable(old) == comparable(offer):
            return []
        changes = []
        for path, label, kind in self.tracked:
            a, b = dig(old, path), dig(offer, path)
            if a != b:
                changes.append(f"{label} : {fmt(a, kind) or '—'} → {fmt(b, kind) or '—'}")
        rec["liste"] = offer
        self.touch(rec["guid"])
        if changes:
            self.event(when, "Modification", rec, " | ".join(changes))
            if rec.get("fiche"):
                rec["fiche_statut"] = "a_rafraichir"  # la fiche détaillée sera relue au prochain passage
        return changes
