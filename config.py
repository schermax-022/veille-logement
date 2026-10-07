"""Réglages, lus dans les variables d'environnement (ou un fichier .env à côté du script)."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def _load_dotenv(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        value = re.split(r"\s+#", value, maxsplit=1)[0].strip().strip('"').strip("'")
        os.environ.setdefault(key.strip(), value)


_load_dotenv(ROOT / ".env")


def env(name: str, default: str = "") -> str:
    value = os.environ.get(name, "").strip()
    return value or default


def env_bool(name: str, default: bool = False) -> bool:
    return env(name, "true" if default else "false").lower() in ("1", "true", "yes", "oui", "on")


# --- Site et API ------------------------------------------------------------
SITE_BASE = env("SITE_BASE", "https://logement-actionlogement.fr").rstrip("/")
API_BASE = env("API_BASE", "https://api.logement-actionlogement.fr/api/v1").rstrip("/")
DETAIL_URL = SITE_BASE + "/search/detail/{guid}"

# Recherche envoyée à l'API : par défaut toute la France, sans filtre.
SEARCH_BODY = json.loads(env("SEARCH_BODY", json.dumps(
    {"municipalities": [], "searchRadiusInKm": 0, "typologyCodes": [], "productGuids": []})))
PAGE_SIZE = int(env("PAGE_SIZE", "100"))           # ramené à 20 automatiquement si refusé
REQUEST_DELAY = float(env("REQUEST_DELAY", "0.5"))  # pause entre deux appels (politesse)
MAX_PAGES = int(env("MAX_PAGES", "500"))
MIN_COMPLETENESS = float(env("MIN_COMPLETENESS", "0.95"))
USE_BROWSER = env_bool("USE_BROWSER")               # passer par un vrai navigateur si l'API bloque

# --- Suivi --------------------------------------------------------------------
MAX_DETAILS_PER_RUN = int(env("MAX_DETAILS_PER_RUN", "300"))  # fiches détaillées récupérées par passage
REMOVAL_CONFIRM_RUNS = int(env("REMOVAL_CONFIRM_RUNS", "2"))  # absences d'affilée avant de valider un retrait
EMPTY_ALERT_AFTER = int(env("EMPTY_ALERT_AFTER", "6"))

DATA_DIR = Path(env("DATA_DIR", str(ROOT / "data")))
EXPORT_PATH = Path(env("EXPORT_PATH", str(ROOT / "export" / "offres-action-logement.xlsx")))

# --- Zone des alertes ---------------------------------------------------------
# Codes INSEE des communes (pas les codes postaux). « 751* » = les 20 arrondissements de Paris.
DEFAULT_ZONES = {
    "751*": "Paris (tous arrondissements)",
    "75056": "Paris",
    "94080": "Vincennes",
    "94067": "Saint-Mandé",
    "92012": "Boulogne-Billancourt",
    "92040": "Issy-les-Moulineaux",
    "92051": "Neuilly-sur-Seine",
    "92044": "Levallois-Perret",
    "94018": "Charenton-le-Pont",
    "93048": "Montreuil",
}
ALERT_ZONES = [z.strip() for z in env("ALERT_ZONES", ",".join(DEFAULT_ZONES)).split(",") if z.strip()]


def in_alert_zone(insee_code: str | None) -> bool:
    code = (insee_code or "").strip()
    if not code:
        return False
    for zone in ALERT_ZONES:
        if zone.endswith("*") and code.startswith(zone[:-1]):
            return True
        if code == zone:
            return True
    return False


ALERT_ON_FIRST_RUN = env_bool("ALERT_ON_FIRST_RUN")
ALERT_ON_RELISTING = env_bool("ALERT_ON_RELISTING", True)  # alerter si une offre retirée revient
MAX_ALERTS_PER_RUN = int(env("MAX_ALERTS_PER_RUN", "25"))
MAX_PHOTOS = int(env("MAX_PHOTOS", "4"))

# --- E-mail -------------------------------------------------------------------
SMTP_HOST = env("SMTP_HOST")
SMTP_PORT = int(env("SMTP_PORT", "465"))
SMTP_SECURITY = env("SMTP_SECURITY", "ssl" if SMTP_PORT == 465 else "starttls").lower()
SMTP_USER = env("SMTP_USER")
SMTP_PASSWORD = env("SMTP_PASSWORD")
EMAIL_FROM = env("EMAIL_FROM", SMTP_USER)
EMAIL_TO = [a.strip() for a in env("EMAIL_TO").split(",") if a.strip()]
EMAIL_ENABLED = bool(SMTP_HOST and EMAIL_TO)

USER_AGENT = env("USER_AGENT", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                               "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36")
