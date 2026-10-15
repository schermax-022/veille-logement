"""Second profil d'alerte : GrandAngoulême.

Les offres des 38 communes de la communauté d'agglomération GrandAngoulême, à 600 € par mois maximum,
sont envoyées à une autre adresse (secret EMAIL_TO_ANGOULEME). Sans ce secret, le profil est inactif.

Action Logement donne le code INSEE de la commune (comparaison exacte). In'li et CDC Habitat ne donnent que
le nom et le code postal : on compare le nom de la commune, avec un code postal de Charente (16…), pour
éviter les homonymes (Brie, Claix, Marsac… existent ailleurs).
"""

from __future__ import annotations

import re
import unicodedata

import config

NAME = "GrandAngoulême"

# Liste Insee de l'intercommunalité 200071827 (CA du Grand Angoulême), 38 communes
COMMUNES = {
    "16015": "Angoulême", "16019": "Asnières-sur-Nouère", "16026": "Balzac", "16055": "Bouëx", "16061": "Brie",
    "16078": "Champniers", "16101": "Claix", "16113": "La Couronne", "16119": "Dignac", "16120": "Dirac",
    "16138": "Fléac", "16146": "Garat", "16154": "Gond-Pontouvre", "16166": "L'Isle-d'Espagnac",
    "16168": "Jauldes", "16187": "Linars", "16199": "Magnac-sur-Touvre", "16210": "Marsac", "16232": "Mornac",
    "16236": "Mouthiers-sur-Boëme", "16244": "Nersac", "16263": "Plassac-Rouffiac", "16271": "Puymoyen",
    "16287": "Roullet-Saint-Estèphe", "16291": "Ruelle-sur-Touvre", "16341": "Saint-Michel",
    "16348": "Saint-Saturnin", "16358": "Saint-Yrieix-sur-Charente", "16368": "Sers", "16370": "Sireuil",
    "16374": "Soyaux", "16382": "Torsac", "16385": "Touvre", "16388": "Trois-Palis", "16415": "Vindelle",
    "16418": "Vœuil-et-Giget", "16420": "Voulgézac", "16422": "Vouzan",
}

TO = [a.strip() for a in config.env("EMAIL_TO_ANGOULEME").split(",") if a.strip()]
ENABLED = bool(TO)
_max = config.env("ANGOULEME_MAX_RENT", "600").replace(",", ".").replace("€", "").strip()
MAX_RENT = float(_max) if _max not in ("", "0") else None


def norm(name: str | None) -> str:
    """« ST YRIEIX SUR CHARENTE », « Saint-Yrieix-sur-Charente » → « SAINT YRIEIX SUR CHARENTE »."""
    t = (name or "").replace("œ", "oe").replace("Œ", "OE")
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode().upper()
    t = re.sub(r"[^A-Z0-9]+", " ", t).strip()
    t = re.sub(r"\bSTE\b", "SAINTE", re.sub(r"\bST\b", "SAINT", t))
    return t


NAMES = {norm(n) for n in COMMUNES.values()}


def by_insee(code: str | None) -> bool:
    return ENABLED and str(code or "") in COMMUNES


def by_place(postcode: str | None, city: str | None) -> bool:
    return ENABLED and str(postcode or "").startswith("16") and norm(city) in NAMES


def rent_ok(rent: float | None) -> bool:
    return MAX_RENT is None or rent is None or rent <= MAX_RENT
