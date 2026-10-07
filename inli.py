"""Veille In'li (www.inli.fr) : même principe que pour Action Logement.

Ici, pas d'API : les annonces sont écrites directement dans les pages du site.
  /locations/offres/?page=N                   24 vignettes par page (+ données de la carte : adresse, GPS)
  /location-<type>-<ville>-<cp>/<référence>   fiche détaillée (loyers, charges, dépôt, DPE, photos…)

Les annonces sont identifiées par leur référence (ex. PRV-251606), dernière partie du lien.
"""

from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup

import config
from al_api import log
from annonce import Field, fmt, parse_dt
from stockage import Store, clean, now_iso
from suivi import bilan, detail_queue, handle_failure, pending_alerts, reconcile

SITE = "In'li"
BASE = config.env("INLI_BASE", "https://www.inli.fr").rstrip("/")
LIST_PATH = "/locations/offres/"
DATA_DIR = config.DATA_DIR / "inli"
EXPORT_PATH = Path(config.env("INLI_EXPORT_PATH", str(config.EXPORT_PATH.parent / "offres-inli.xlsx")))
ENABLED = config.env_bool("INLI_ENABLED", True)
MAX_DETAILS_PER_RUN = int(config.env("INLI_MAX_DETAILS_PER_RUN", "150"))
MIN_EXPECTED = int(config.env("INLI_MIN_EXPECTED", "50"))  # moins d'annonces que ça = page probablement bloquée

# Zone d'alerte par code postal (In'li ne donne pas le code INSEE). « 75* » = tout Paris.
# Par défaut : Paris, Vincennes, Saint-Mandé, Boulogne-Billancourt, Issy-les-Moulineaux,
# Neuilly-sur-Seine, Levallois-Perret, Charenton-le-Pont, Montreuil.
ZONES = [z.strip() for z in config.env(
    "INLI_ALERT_POSTCODES", "75*,94300,94160,92100,92130,92200,92300,94220,93100").split(",") if z.strip()]
_max = config.env("INLI_ALERT_MAX_RENT", config.env("ALERT_MAX_RENT", "1100")).replace(",", ".").replace("€", "")
MAX_RENT = float(_max) if _max.strip() not in ("", "0") else None   # loyer charges comprises maximum

TRACKED = [("loyer_cc", "Loyer cc", "money"), ("surface", "Surface", "area"), ("pieces", "Pièces", "text")]


class InliError(Exception):
    pass


# -----------------------------------------------------------------------------------------
# Accès au site (direct, ou via un navigateur si le site renvoie des pages vides aux robots)
# -----------------------------------------------------------------------------------------

class HtmlClient:
    def __init__(self) -> None:
        self.mode = "direct"
        self.calls = 0
        self._pw = self._browser = self._ctx = self._page = None

    def _direct(self, url: str) -> tuple[int, str]:
        req = urllib.request.Request(url, headers={
            "User-Agent": config.USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "fr-FR,fr;q=0.9",
        })
        try:
            with urllib.request.urlopen(req, timeout=40) as resp:
                return resp.status, resp.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as err:
            return err.code, err.read().decode("utf-8", "replace")

    def start_browser(self) -> bool:
        try:
            from playwright.sync_api import sync_playwright
            self._pw = sync_playwright().start()
            self._browser = self._pw.chromium.launch(headless=True)
            self._ctx = self._browser.new_context(locale="fr-FR", user_agent=config.USER_AGENT)
            # inutile de charger photos, polices et vidéos pour lire les annonces : plus rapide
            self._ctx.route("**/*", lambda route: route.abort()
                            if route.request.resource_type in ("image", "font", "media") else route.continue_())
            self._page = self._ctx.new_page()
        except Exception as exc:
            log(f"⚠️  {SITE} : navigateur indisponible ({type(exc).__name__}). "
                "Sur GitHub, ajoutez la variable USE_BROWSER = true.")
            return False
        self.mode = "browser"
        log(f"{SITE} : passage en mode navigateur.")
        return True

    def get(self, url: str, attempts: int = 4) -> str:
        last = "?"
        for attempt in range(attempts):
            if attempt:
                time.sleep(min(30, 2 ** attempt))
            self.calls += 1
            try:
                if self.mode == "browser":
                    resp = self._page.goto(url, wait_until="domcontentloaded", timeout=60000)
                    status, text = (resp.status if resp else 0), self._page.content()
                else:
                    status, text = self._direct(url)
            except Exception as exc:
                last = f"{type(exc).__name__}: {exc}"
                continue
            if status == 200:
                return text
            last = f"HTTP {status}"
            if status in (404, 410):
                break
            if status == 429 or status >= 500:
                continue
            break
        raise InliError(f"{url} : {last}")

    def get_bytes(self, url: str, max_bytes: int = 6_000_000) -> tuple[bytes, str]:
        if self.mode == "browser":
            resp = self._ctx.request.get(url, timeout=30000)
            if not resp.ok:
                raise InliError(f"HTTP {resp.status}")
            return resp.body()[:max_bytes], resp.headers.get("content-type", "")
        req = urllib.request.Request(url, headers={"User-Agent": config.USER_AGENT, "Referer": BASE + "/"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read(max_bytes + 1)[:max_bytes], resp.headers.get("Content-Type", "")

    def close(self) -> None:
        for obj in (self._browser, self._pw):
            try:
                if obj is self._pw and obj:
                    obj.stop()
                elif obj:
                    obj.close()
            except Exception:
                pass


# -----------------------------------------------------------------------------------------
# Lecture des pages
# -----------------------------------------------------------------------------------------

def money(text: Any) -> float | None:
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    t = re.sub(r"[\s \xa0]", "", str(text))
    m = re.search(r"\d+(?:[.,]\d+)?", t)
    return float(m.group(0).replace(",", ".")) if m else None


def squash(text: str) -> str:
    return re.sub(r"[ \t \xa0]+", " ", text or "").strip()


def ref_of(link: str) -> str:
    return link.rstrip("/").rsplit("/", 1)[-1]


def _markers(html: str) -> dict[str, dict]:
    m = re.search(r"window\.ALENTOOR_LISTING\s*=\s*\{\s*markers:\s*(\[.*?\])\s*\}\s*;", html, re.S)
    if not m:
        return {}
    try:
        groups = json.loads(m.group(1))
    except json.JSONDecodeError:
        return {}
    items = [it for g in groups for it in (g.get("items") or [g])]
    return {it["link"]: it for it in items if it.get("link")}


def parse_list_page(html: str) -> tuple[int | None, int, dict[str, dict]]:
    """Renvoie (nombre total annoncé, numéro de la dernière page, offres de la page)."""
    main = html.split("Offres pouvant vous intéresser")[0]   # on ignore les suggestions en bas de page
    soup = BeautifulSoup(main, "html.parser")
    total = None
    for h5 in soup.find_all("h5"):
        m = re.fullmatch(r"([\d\s \xa0]+)\s+logements?", squash(h5.get_text(" ")))
        if m:
            total = int(re.sub(r"\D", "", m.group(1)))
            break
    if total is None and "Aucune annonce" in main:
        total = 0
    pages = [int(n) for n in re.findall(r'href="/locations/offres/[^"]*[?&]page=(\d+)"', html)]
    last_page = max(pages) if pages else 1

    markers = _markers(html)
    offers: dict[str, dict] = {}
    for price in soup.select(".featured-price"):
        a = price.find_parent("a")
        if not a or not a.get("href"):
            continue
        href = a["href"]
        ref = ref_of(href)
        lines = [squash(x) for x in a.select_one(".featured-details").get_text("\n").split("\n")] \
            if a.select_one(".featured-details") else []
        lines = [x for x in lines if x]
        badge = a.select_one(".featured-badge")
        img = a.find("img")
        mk = markers.get(href, {})
        cp = mk.get("postalCode") or (re.search(r"-(\d{5})/", href) or [None, None])[1]
        kind = (re.match(r"/location-([a-z]+)-", href) or [None, None])[1]
        summary = lines[0] if lines else ""
        offer = {
            "guid": ref,
            "ref": ref,
            "lien": href,
            "url": BASE + href,
            "loyer_cc": money(mk.get("price")) if mk.get("price") else money(price.get_text(" ")),
            "resume": summary,
            "ville_affichee": lines[1] if len(lines) > 1 else None,
            "type": mk.get("type") or (kind.capitalize() if kind else None),
            "pieces": mk.get("room") or (int(re.search(r"(\d+)\s*pi", summary).group(1))
                                         if re.search(r"(\d+)\s*pi", summary) else None),
            "surface": mk.get("area") or money((re.search(r"([\d.,]+)\s*m²", summary) or [None, None])[1]),
            "adresse": mk.get("address"),
            "ville": mk.get("city") or (lines[1] if len(lines) > 1 else None),
            "code_postal": cp,
            "latitude": mk.get("lat"),
            "longitude": mk.get("lng"),
            "badge": squash(badge.get_text(" ")) if badge else None,
            "photo": mk.get("image") or (img.get("src") if img else None),
            "programme": mk.get("id"),
        }
        offers[ref] = {k: v for k, v in offer.items() if v not in (None, "")}
    return total, last_page, offers


def parse_detail(html: str) -> dict | None:
    soup = BeautifulSoup(html, "html.parser")
    block = soup.select_one(".advert-price-block")
    if not block:
        return None
    d: dict[str, Any] = {}
    title = soup.select_one(".advert-page-title")
    if title:
        for key, sel in (("titre", "h1"), ("resume", "h3"), ("lieu", "h4"), ("reference_affichee", ".reference")):
            el = title.select_one(sel)
            if el:
                d[key] = squash(el.get_text(" "))
        if d.get("reference_affichee"):
            d["reference_affichee"] = re.sub(r"^Référence logement\s*:\s*", "", d["reference_affichee"])
    badge = block.select_one(".badge")
    if badge:
        d["statut_annonce"] = squash(badge.get_text(" "))
    d["avantages"] = [squash(x.get_text(" ")) for x in block.select(".advert-perk-inline")]
    h5 = block.select_one("h5")
    if h5:
        d["loyer_cc_affiche"] = money(h5.get_text(" "))
    lines = []
    for li in block.select("ul > li"):
        span = li.find("span")
        value = squash(span.get_text(" ")) if span else None
        if span:
            span.extract()
        lines.append({"libelle": squash(li.get_text(" ")), "valeur": value})
    d["lignes_prix"] = lines
    sections = {}
    for sec in soup.select(".advert-summary-section"):
        head = sec.find("h5")
        items = [squash(x.get_text(" ")) for x in sec.select(".advert-summary-item")]
        if head:
            sections[squash(head.get_text(" "))] = [i for i in items if i]
    d["caracteristiques"] = sections
    desc = soup.select_one(".advert-description-content")
    if desc:
        paras = [squash(p.get_text(" ")) for p in desc.find_all(["p", "li"])] or [squash(desc.get_text(" "))]
        d["description"] = "\n".join(p for p in paras if p)
    for key, prefix in (("dpe", "consumption-"), ("ges", "gaz-effect-")):
        el = soup.find(lambda t: t.name and t.get("data-step") and "active" in (t.get("class") or [])
                       and any(c.startswith(prefix) for c in (t.get("class") or [])))
        if el:
            d[key] = {"classe": el.get("data-step"), "valeur": money(el.get("data-value"))}
    photos = []
    for img in soup.select(".advert-carousel img"):
        src = img.get("src") or img.get("data-src")
        if src and src not in photos:
            photos.append(src)
    d["photos"] = photos
    rented = soup.select_one(".advert-rented-by")
    if rented:
        txt = squash(rented.get_text(" ").replace("Ce bien est loué par", ""))
        if txt:
            d["loue_par"] = txt
    return d


@dataclass
class Crawl:
    offers: dict = field(default_factory=dict)
    total: int = 0
    pages: int = 0
    errors: list = field(default_factory=list)

    @property
    def complete(self) -> bool:
        return not self.errors and self.total > 0 and len(self.offers) >= config.MIN_COMPLETENESS * self.total


def fetch_all(client: HtmlClient, expected_min: int) -> Crawl:
    url = BASE + LIST_PATH + "?page={}"
    html = client.get(url.format(1))
    total, last_page, offers = parse_list_page(html)
    if (total or 0) < expected_min and client.mode == "direct":
        log(f"⚠️  {SITE} : la page annonce {total} logement(s), bien moins que prévu (≥ {expected_min}) : "
            "le site renvoie peut-être une page allégée aux robots.")
        if client.start_browser():
            html = client.get(url.format(1))
            total, last_page, offers = parse_list_page(html)
    if (total or 0) < expected_min:
        raise InliError(f"page incomplète : {total} logement(s) annoncé(s), au moins {expected_min} attendus")
    crawl = Crawl(offers=dict(offers), total=total, pages=last_page)
    for page in range(2, min(last_page, config.MAX_PAGES) + 1):
        time.sleep(config.REQUEST_DELAY)
        try:
            _, _, offers = parse_list_page(client.get(url.format(page)))
        except InliError as exc:
            crawl.errors.append(f"page {page} : {exc}")
            continue
        if not offers:
            crawl.errors.append(f"page {page} vide")
        crawl.offers.update(offers)
    return crawl


def fetch_detail(client: HtmlClient, url: str) -> tuple[dict | None, str]:
    try:
        data = parse_detail(client.get(url, attempts=3))
    except InliError as exc:
        return None, f"erreur : {exc}"
    return (data, "ok") if data else (None, "fiche introuvable")


# -----------------------------------------------------------------------------------------
# Champs (Excel + e-mail)
# -----------------------------------------------------------------------------------------

def L(rec: dict, key: str) -> Any:
    return (rec.get("liste") or {}).get(key)


def Fi(rec: dict, key: str) -> Any:
    return (rec.get("fiche") or {}).get(key)


KNOWN_LINES = {
    "loyer_hc": r"^loyer hors charges",
    "charges": r"^provisions? (de|sur) charges",
    "parking": r"^loyer.*stationnement",
    "depot": r"^d[ée]p[ôo]t de garantie",
    "ref_majore": r"^loyer de r[ée]f[ée]rence major[ée]",
    "ce_bien": r"^loyer de ce bien",
    "encadrement": r"encadrement des loyers",
}


def price_line(rec: dict, key: str) -> dict | None:
    for line in Fi(rec, "lignes_prix") or []:
        if re.search(KNOWN_LINES[key], line.get("libelle") or "", re.I):
            return line
    return None


def price_value(rec: dict, key: str) -> float | None:
    line = price_line(rec, key)
    return money(line.get("valeur")) if line else None


def other_lines(rec: dict) -> str | None:
    out = []
    for line in Fi(rec, "lignes_prix") or []:
        lab = line.get("libelle") or ""
        if any(re.search(p, lab, re.I) for p in KNOWN_LINES.values()):
            continue
        out.append(f"{lab} : {line['valeur']}" if line.get("valeur") else lab)
    return " ; ".join(out) or None


def essentials(rec: dict) -> list[str]:
    return (Fi(rec, "caracteristiques") or {}).get("L'essentiel") or []


def rooms(rec: dict) -> Any:
    for item in essentials(rec):
        m = re.match(r"(\d+)\s*pi[eè]ce", item, re.I)
        if m:
            return int(m.group(1))
        if re.match(r"studio", item, re.I):
            return 1
    return L(rec, "pieces")


def bedrooms(rec: dict) -> int | None:
    for item in essentials(rec):
        m = re.match(r"(\d+)\s*chambre", item, re.I)
        if m:
            return int(m.group(1))
    return None


def area(rec: dict) -> float | None:
    for item in essentials(rec):
        if "m²" in item or "m2" in item:
            return money(item)
    return L(rec, "surface")


def floor(rec: dict) -> str | None:
    for item in essentials(rec):
        if not re.search(r"pi[eè]ce|chambre|m²|m2|studio", item, re.I):
            return item
    return None


def section(rec: dict, name: str) -> str | None:
    items = (Fi(rec, "caracteristiques") or {}).get(name)
    return " ; ".join(items) if items else None


def other_sections(rec: dict) -> str | None:
    secs = Fi(rec, "caracteristiques") or {}
    out = [f"{k} : {', '.join(v)}" for k, v in secs.items() if k not in ("L'essentiel", "Les plus", "La résidence") and v]
    return " ; ".join(out) or None


def whole(value: Any) -> Any:
    return int(value) if isinstance(value, float) and value.is_integer() else value


def rent_cc(rec: dict) -> float | None:
    return L(rec, "loyer_cc") or Fi(rec, "loyer_cc_affiche")


def in_zone_cp(cp: str | None) -> bool:
    cp = (cp or "").strip()
    return bool(cp) and any((cp.startswith(z[:-1]) if z.endswith("*") else cp == z) for z in ZONES)


def is_in_zone(rec: dict) -> bool:
    return in_zone_cp(L(rec, "code_postal"))


def offer_in_zone(offer: dict) -> bool:
    return in_zone_cp(offer.get("code_postal"))


def wants_alert(offer: dict) -> bool:
    if not offer_in_zone(offer):
        return False
    rent = offer.get("loyer_cc")
    return MAX_RENT is None or rent is None or rent <= MAX_RENT


def city_label(rec: dict) -> str:
    shown = L(rec, "ville_affichee") or Fi(rec, "lieu") or L(rec, "ville") or ""
    return shown.replace("ème", "e").replace("Eme", "e").replace("EME", "e")


def describe(offer: dict) -> str:
    return f"{offer.get('resume') or ''} {offer.get('ville_affichee') or offer.get('ville') or ''}".strip()


def summary(rec: dict) -> dict:
    return {"reference": L(rec, "ref") or rec.get("guid", ""), "ville": L(rec, "ville_affichee") or L(rec, "ville") or "",
            "code_postal": L(rec, "code_postal") or "", "typologie": L(rec, "resume") or "",
            "loyer": fmt(rent_cc(rec), "money") or ""}


STATUS_FR = {"en_ligne": "En ligne", "retiree": "Retirée"}


def F(section_, label, get, kind="text", width=14, in_email=True):
    return Field(section_, label, get, kind, (), width, in_email)


FIELDS: list[Field] = [
    F("Suivi", "Statut", lambda r: STATUS_FR.get(r.get("statut"), r.get("statut")), width=11, in_email=False),
    F("Suivi", "Mise en ligne détectée le", lambda r: parse_dt(r.get("premiere_detection")), "datetime", 18, False),
    F("Suivi", "Déjà en ligne au démarrage du suivi", lambda r: r.get("deja_en_ligne_au_demarrage"), "bool", 14, False),
    F("Suivi", "Vue en ligne pour la dernière fois", lambda r: parse_dt(r.get("derniere_presence")), "datetime", 18,
      False),
    F("Suivi", "Retrait détecté le", lambda r: parse_dt(r.get("retrait_detecte")), "datetime", 18, False),
    F("Suivi", "Durée en ligne (jours)", lambda r: None, "formula", 12, False),
    F("Suivi", "Remises en ligne", lambda r: r.get("remises_en_ligne") or 0, "int", 10, False),
    F("Suivi", "Zone d'alerte", is_in_zone, "bool", 10, False),
    F("Localisation", "Ville", lambda r: L(r, "ville_affichee") or L(r, "ville"), width=24),
    F("Localisation", "Code postal", lambda r: L(r, "code_postal"), width=10),
    F("Localisation", "Département", lambda r: (L(r, "code_postal") or "")[:2] or None, width=8, in_email=False),
    F("Localisation", "Adresse", lambda r: L(r, "adresse"), width=28),
    F("Localisation", "Latitude", lambda r: L(r, "latitude"), width=10, in_email=False),
    F("Localisation", "Longitude", lambda r: L(r, "longitude"), width=10, in_email=False),
    F("Loyer et frais", "Loyer charges comprises (€/mois)", rent_cc, "money", 13),
    F("Loyer et frais", "Loyer hors charges (€/mois)", lambda r: price_value(r, "loyer_hc"), "money", 12),
    F("Loyer et frais", "Provisions de charges (€/mois)", lambda r: price_value(r, "charges"), "money", 12),
    F("Loyer et frais", "Stationnement charges comprises (€/mois)", lambda r: price_value(r, "parking"), "money", 13),
    F("Loyer et frais", "Dépôt de garantie (€)", lambda r: price_value(r, "depot"), "money", 12),
    F("Loyer et frais", "Encadrement des loyers", lambda r: bool(price_line(r, "encadrement")) if r.get("fiche") else None,
      "bool", 11),
    F("Loyer et frais", "Loyer de référence majoré (€/m² hors charges)", lambda r: price_value(r, "ref_majore"), "money",
      13),
    F("Loyer et frais", "Avantages", lambda r: " ; ".join(Fi(r, "avantages") or []) or None, width=18),
    F("Loyer et frais", "Autres montants", other_lines, width=24),
    F("Logement", "Type", lambda r: L(r, "type"), width=12),
    F("Logement", "Mention", lambda r: L(r, "badge"), width=10),
    F("Logement", "Pièces", rooms, "int", 7),
    F("Logement", "Chambres", bedrooms, "int", 8),
    F("Logement", "Surface (m²)", area, "area", 9),
    F("Logement", "Étage", floor, width=9),
    F("Logement", "Les plus", lambda r: section(r, "Les plus"), width=22),
    F("Logement", "La résidence", lambda r: section(r, "La résidence"), width=20),
    F("Logement", "Autres caractéristiques", other_sections, width=22),
    F("Énergie", "DPE (classe)", lambda r: (Fi(r, "dpe") or {}).get("classe"), width=8),
    F("Énergie", "Consommation (kWh/m²/an)", lambda r: whole((Fi(r, "dpe") or {}).get("valeur")), "int", 11),
    F("Énergie", "GES (classe)", lambda r: (Fi(r, "ges") or {}).get("classe"), width=8),
    F("Énergie", "Émissions (kg CO₂/m²/an)", lambda r: whole((Fi(r, "ges") or {}).get("valeur")), "int", 11),
    F("Description", "Titre", lambda r: Fi(r, "titre") or L(r, "resume"), width=30),
    F("Description", "Description", lambda r: Fi(r, "description"), "longtext", 60),
    F("Description", "Photos", lambda r: len(Fi(r, "photos") or []) or (1 if L(r, "photo") else 0), "int", 7, False),
    F("Références", "Lien vers l'annonce", lambda r: L(r, "url"), "link", 14, False),
    F("Références", "Référence", lambda r: L(r, "ref") or r.get("guid"), width=14),
    F("Références", "Loué par", lambda r: Fi(r, "loue_par"), width=16),
    F("Références", "Programme (identifiant carte)", lambda r: L(r, "programme"), width=10, in_email=False),
]


def profile(store: Store):
    import export_excel
    urls = {g: L(r, "url") for g, r in store.records.items()}
    return export_excel.Profile(
        title="Suivi des offres — inli.fr", fields=FIELDS, extra_fn=lambda r: {}, zone_fn=is_in_zone,
        detail_url=lambda guid: urls.get(guid) or BASE + LIST_PATH,
        notes=[("Zone d'alerte (codes postaux)", ", ".join(ZONES)),
               ("Plafond des alertes", f"{fmt(MAX_RENT, 'money')} par mois charges comprises" if MAX_RENT else "aucun"),
               ("Source", f"Pages du site {BASE}{LIST_PATH} (24 annonces par page) et fiches détaillées")])


# -----------------------------------------------------------------------------------------
# E-mail
# -----------------------------------------------------------------------------------------

def render(rec: dict, cids: list[str], relisted: bool = False) -> tuple[str, str, str]:
    from alertes import ACCENT, INK, LINE, MUTED, h

    rent = rent_cc(rec)
    city = city_label(rec)
    bits = [f"{rooms(rec)} pièces" if rooms(rec) else None, fmt(area(rec), "area"),
            f"{fmt(rent, 'money')}/mois cc" if rent else None]
    prefix = "🔁 In'li · de nouveau en ligne" if relisted else "🏠 In'li · nouvelle annonce"
    subj = f"{prefix} · {city} · " + " · ".join(b for b in bits if b)

    detected = parse_dt(rec.get("premiere_detection"))
    when = f"détectée le {(detected or parse_dt(now_iso())).strftime('%d/%m/%Y à %H:%M')}"
    kicker = "In'li · De nouveau en ligne" if relisted else "In'li · Nouvelle annonce"
    head_title = Fi(rec, "titre") or f"{L(rec, 'type') or 'Logement'} à {city}"
    address = ", ".join(x for x in (L(rec, "adresse"), f"{L(rec, 'code_postal') or ''} {L(rec, 'ville') or ''}".strip()) if x)
    url = L(rec, "url")
    by_label = {f.label: fmt(f.get(rec), f.kind) for f in FIELDS}
    facts = [(k, by_label.get(k)) for k in ("Pièces", "Surface (m²)", "Étage", "Chambres")]
    facts = [(k.replace(" (m²)", ""), v) for k, v in facts if v]
    groups = [
        ("Loyer et frais", ["Loyer hors charges (€/mois)", "Provisions de charges (€/mois)",
                            "Stationnement charges comprises (€/mois)", "Dépôt de garantie (€)", "Encadrement des loyers",
                            "Loyer de référence majoré (€/m² hors charges)", "Avantages", "Autres montants"]),
        ("Le logement", ["Type", "Mention", "Les plus", "La résidence", "Autres caractéristiques"]),
        ("Énergie", ["DPE (classe)", "Consommation (kWh/m²/an)", "GES (classe)", "Émissions (kg CO₂/m²/an)"]),
        ("Références", ["Référence", "Loué par"]),
    ]
    short = lambda k: k.replace(" (€/mois)", "").replace(" (€)", "")
    sections = [(t, [(short(k), by_label[k]) for k in keys if by_label.get(k)]) for t, keys in groups]
    sections = [(t, rows) for t, rows in sections if rows]
    description = Fi(rec, "description")

    lines = [f"{kicker} — {when}", "", head_title]
    if address:
        lines.append(address)
    if rent:
        lines.append(f"Loyer : {fmt(rent, 'money')} / mois charges comprises")
    lines += [f"{k} : {v}" for k, v in facts]
    for t, rows in sections:
        lines += ["", t.upper()] + [f"  {k} : {v}" for k, v in rows]
    if description:
        lines += ["", "DESCRIPTION", description]
    lines += ["", f"Voir l'annonce : {url}"]
    text = "\n".join(lines) + "\n"

    def table(rows):
        return "".join(
            f'<tr><td style="padding:7px 0;border-top:1px solid {LINE};color:{MUTED};font-size:13px;width:44%;'
            f'vertical-align:top">{h(k)}</td><td style="padding:7px 0;border-top:1px solid {LINE};color:{INK};'
            f'font-size:14px;vertical-align:top">{h(v)}</td></tr>' for k, v in rows)

    photos_html = ""
    if cids:
        photos_html = (f'<tr><td style="padding:0 0 6px 0"><img src="cid:{cids[0]}" width="560" alt="Photo 1" '
                       f'style="display:block;width:100%;max-width:560px;height:auto;border-radius:8px"></td></tr>')
        thumbs = cids[1:4]
        if thumbs:
            cells = "".join(
                f'<td width="{100 // len(thumbs)}%" style="padding:0 {3 if j < len(thumbs) - 1 else 0}px 0 0">'
                f'<img src="cid:{c}" alt="Photo {j + 2}" style="display:block;width:100%;height:auto;border-radius:6px">'
                f'</td>' for j, c in enumerate(thumbs))
            photos_html += (f'<tr><td style="padding:0 0 8px 0"><table role="presentation" width="100%" '
                            f'cellpadding="0" cellspacing="0"><tr>{cells}</tr></table></td></tr>')
    facts_html = "".join(
        f'<td style="padding:10px 12px;background:#f6f7f9;border-radius:8px;text-align:center">'
        f'<div style="font-size:11px;color:{MUTED};text-transform:uppercase;letter-spacing:.04em">{h(k)}</div>'
        f'<div style="font-size:15px;font-weight:600;color:{INK};margin-top:2px">{h(v)}</div></td>'
        f'<td style="width:6px"></td>' for k, v in facts[:4])

    def block(t, rows):
        return (f'<tr><td style="padding:22px 0 4px 0;font-size:15px;font-weight:700;color:{INK}">{h(t)}</td></tr>'
                f'<tr><td><table role="presentation" width="100%" cellpadding="0" cellspacing="0">{table(rows)}'
                f'</table></td></tr>')

    top = "".join(block(t, rows) for t, rows in sections if t in ("Loyer et frais", "Le logement"))
    bottom = "".join(block(t, rows) for t, rows in sections if t not in ("Loyer et frais", "Le logement"))
    desc_html = (f'<tr><td style="padding:22px 0 6px 0;font-size:15px;font-weight:700;color:{INK}">Description</td></tr>'
                 f'<tr><td style="font-size:14px;line-height:1.55;color:{INK};white-space:pre-line">{h(description)}'
                 f'</td></tr>') if description else ""
    no_detail = ('<tr><td style="padding:10px 12px;background:#fff6e5;border-radius:8px;font-size:13px;color:#7a4b00">'
                 'La fiche détaillée n\'a pas pu être lue : seules les informations de la liste sont affichées.'
                 '</td></tr>') if not rec.get("fiche") else ""
    rent_html = (f'<tr><td style="padding:10px 0 2px 0"><span style="font-size:26px;font-weight:800;color:{INK}">'
                 f'{h(fmt(rent, "money"))}</span> <span style="font-size:14px;color:{MUTED}">/ mois charges comprises'
                 f'</span></td></tr>') if rent else ""
    html = f"""<!doctype html><html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eef0f3">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef0f3">
<tr><td align="center" style="padding:20px 10px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
 style="max-width:600px;background:#ffffff;border-radius:12px;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif">
<tr><td style="padding:22px 20px 26px 20px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0">
  <tr><td style="font-size:12px;color:{ACCENT};font-weight:700;text-transform:uppercase;letter-spacing:.05em">
    {h(kicker)} <span style="color:{MUTED};font-weight:400;text-transform:none;letter-spacing:0">· {h(when)}</span></td></tr>
  <tr><td style="padding:6px 0 4px 0;font-size:21px;line-height:1.3;font-weight:700;color:{INK}">{h(head_title)}</td></tr>
  <tr><td style="padding:0 0 14px 0;font-size:14px;color:{MUTED}">{h(address or city)}</td></tr>
  {photos_html}
  {rent_html}
  <tr><td style="padding:12px 0 0 0"><table role="presentation" cellpadding="0" cellspacing="0"><tr>{facts_html}</tr></table></td></tr>
  <tr><td style="padding:20px 0 0 0"><a href="{h(url)}" style="display:inline-block;padding:12px 20px;background:{ACCENT};
    color:#ffffff;text-decoration:none;border-radius:8px;font-size:15px;font-weight:700">Voir l'annonce sur inli.fr</a></td></tr>
  {no_detail}
  {top}
  {desc_html}
  {bottom}
  <tr><td style="padding:26px 0 0 0;font-size:12px;color:{MUTED}">
    Alerte envoyée par votre veille In'li · <a href="{h(url)}" style="color:{MUTED}">{h(url)}</a></td></tr>
</table></td></tr></table></td></tr></table></body></html>"""
    return subj, text, html


def photo_urls(rec: dict) -> list[str]:
    return list(Fi(rec, "photos") or []) or ([L(rec, "photo")] if L(rec, "photo") else [])


def send_alert(client: HtmlClient, rec: dict, relisted: bool = False) -> None:
    import alertes
    photos = alertes.download_photos(client, photo_urls(rec))
    alertes.send_rendered(lambda cids: render(rec, cids, relisted), photos)


# -----------------------------------------------------------------------------------------
# Passage
# -----------------------------------------------------------------------------------------

def open_store() -> Store:
    return Store(DATA_DIR, summary=summary, tracked=TRACKED)


def export(store: Store) -> None:
    import export_excel
    export_excel.build(store, EXPORT_PATH, profile(store))
    log(f"Excel In'li mis à jour : {EXPORT_PATH}")


def run() -> int:
    if not ENABLED:
        return 0
    store = open_store()
    client = HtmlClient()
    exit_code = 0
    changed = False
    try:
        now = now_iso()
        last_total = store.state.get("dernier_total")
        expected = max(20, int(0.5 * last_total)) if last_total else MIN_EXPECTED
        try:
            crawl = fetch_all(client, expected)
        except InliError as exc:
            return handle_failure(store, str(exc), SITE, BASE + LIST_PATH)
        log(f"{SITE} : {len(crawl.offers)} offres récupérées sur {crawl.total} annoncées "
            f"({crawl.pages} page(s), mode {client.mode}).")
        for err in crawl.errors:
            log(f"⚠️  {SITE} : {err}")

        counts = reconcile(store, crawl.offers, now, complete=crawl.complete, total=crawl.total,
                           in_zone=offer_in_zone, wants_alert=wants_alert, describe=describe,
                           rent_of=lambda o: o.get("loyer_cc"), max_rent=MAX_RENT)
        if crawl.complete:
            store.set_state(dernier_total=crawl.total)

        # fiches détaillées
        queue, budget = detail_queue(store, is_in_zone, MAX_DETAILS_PER_RUN)
        ok = failures = 0
        for rec in queue[:budget]:
            time.sleep(config.REQUEST_DELAY)
            data, status = fetch_detail(client, L(rec, "url"))
            failures = 0 if data else failures + 1
            if data:
                rec.update(fiche=clean(data), fiche_statut="ok", fiche_recuperee_le=now, fiche_essais=0)
                ok += 1
            else:
                rec.update(fiche_statut=status, fiche_essais=(rec.get("fiche_essais") or 0) + 1)
            store.touch(rec["guid"])
            if failures >= 5:
                log(f"⚠️  {SITE} répond mal : lecture des fiches reportée au prochain passage.")
                break
        if queue:
            log(f"{SITE} — fiches détaillées : {ok} récupérée(s), {len(queue) - ok} en attente.")
        changed |= store.save()

        # alertes
        pending = pending_alerts(store)
        sent = 0
        if pending and not config.EMAIL_ENABLED:
            log(f"⚠️  {len(pending)} alerte(s) {SITE} en attente mais l'e-mail n'est pas configuré.")
        elif pending:
            for rec in pending[: config.MAX_ALERTS_PER_RUN]:
                if not rec.get("fiche"):
                    data, _ = fetch_detail(client, L(rec, "url"))
                    if data:
                        rec.update(fiche=clean(data), fiche_statut="ok", fiche_recuperee_le=now)
                try:
                    send_alert(client, rec, rec["alerte"] == "a_envoyer_remise")
                except Exception as exc:
                    log(f"⚠️  Alerte {SITE} non envoyée ({describe(rec['liste'])}) : {exc} — nouvel essai au prochain passage.")
                    exit_code = 1
                    break
                rec.update(alerte="envoyee", alerte_envoyee_le=now_iso())
                store.touch(rec["guid"])
                sent += 1
                log(f"✉️  Alerte {SITE} envoyée : {describe(rec['liste'])} — {fmt(rent_cc(rec), 'money')}")
            changed |= store.save()

        online = sum(1 for r in store.records.values() if r["statut"] == "en_ligne")
        bilan(SITE, counts, sent, online, len(crawl.offers))
        if changed or not EXPORT_PATH.exists():
            export(store)
        log(f"{SITE} : {client.calls} page(s) consultée(s).")
        return exit_code
    finally:
        client.close()


def run_test_email() -> int:
    """Exemple d'alerte avec une vraie annonce In'li parisienne."""
    import alertes
    client = HtmlClient()
    try:
        _, _, offers = parse_list_page(client.get(BASE + LIST_PATH + "paris_d:75/"))
        if not offers:
            alertes.send_simple("Test — Veille In'li", "Les e-mails fonctionnent 👍 (aucune annonce In'li à Paris "
                                "pour illustrer l'alerte).", BASE + LIST_PATH)
            return 0
        offer = next(iter(offers.values()))
        data, _ = fetch_detail(client, offer["url"])
        rec = {"guid": offer["guid"], "premiere_detection": now_iso(), "liste": offer, "fiche": data}
        send_alert(client, rec)
        log(f"E-mail de test In'li envoyé avec l'annonce : {describe(offer)}")
        return 0
    finally:
        client.close()


def run_export() -> int:
    store = open_store()
    export(store)
    return 0

