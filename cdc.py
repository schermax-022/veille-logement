"""Veille CDC Habitat (www.cdc-habitat.fr, rubrique « Louer ») : même principe que pour les autres sites.

Pas d'API publique, mais la page de résultats contient la liste complète des logements à louer
(toute la France) dans le bloc de données qui alimente la carte :
  /Recherche/show/…cdTypage=Location&     → searchBootstrap.lots : n° de lot, loyer cc, catégorie, GPS
  /Recherche/getInfoBoxLogement (POST)    → la vignette d'un logement : ville, code postal, lien, typologie
  /annonces-immobilieres/location/<région>/<département>/<ville>/<n° de lot>   → fiche détaillée

Une seule page suffit donc à repérer les nouveautés et les retraits ; la vignette (petite) n'est lue
qu'une fois par logement, la fiche détaillée aussi. Les logements sont identifiés par leur n° de lot.
"""

from __future__ import annotations

import http.cookiejar
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup

import config
from al_api import log
from annonce import Field, fmt, parse_dt
from stockage import Store, now_iso
from suivi import bilan, detail_queue, handle_failure, pending_alerts, reconcile
import profils

SITE = "CDC Habitat"
SENDER = "Veille CDC Habitat"
BASE = config.env("CDC_BASE", "https://www.cdc-habitat.fr").rstrip("/")
LIST_PATH = "/Recherche/show/lbLieu=&nbLoyerMax=&cdTypage=Location&"
CARD_PATH = "/Recherche/getInfoBoxLogement"
PUBLIC_LIST = BASE + "/recherche/location"
DATA_DIR = config.DATA_DIR / "cdc"
EXPORT_PATH = Path(config.env("CDC_EXPORT_PATH", str(config.EXPORT_PATH.parent / "offres-cdc-habitat.xlsx")))
ENABLED = config.env_bool("CDC_ENABLED", True)
MAX_DETAILS_PER_RUN = int(config.env("CDC_MAX_DETAILS_PER_RUN", "120"))  # logements complétés par passage
MAX_NEW_CARDS = 150             # vignettes lues avant de décider des alertes (nouveautés uniquement)
MIN_EXPECTED = int(config.env("CDC_MIN_EXPECTED", "100"))  # moins que ça = page probablement bloquée

# Zone d'alerte par code postal. « 75* » = tout Paris. Par défaut : Paris, Vincennes, Saint-Mandé,
# Boulogne-Billancourt, Issy-les-Moulineaux, Neuilly-sur-Seine, Levallois-Perret, Charenton-le-Pont, Montreuil.
ZONES = [z.strip() for z in config.env(
    "CDC_ALERT_POSTCODES", "75*,94300,94160,92100,92130,92200,92300,94220,93100").split(",") if z.strip()]
_max = config.env("CDC_ALERT_MAX_RENT", config.env("ALERT_MAX_RENT", "1100")).replace(",", ".").replace("€", "")
MAX_RENT = float(_max) if _max.strip() not in ("", "0") else None   # loyer charges comprises maximum
# Catégories enregistrées mais jamais signalées par e-mail (les places de parking surtout)
NO_ALERT = {c.strip().lower() for c in config.env("CDC_NO_ALERT_CATEGORIES", "stationnement").split(",") if c.strip()}
# Grand cadre GPS autour de la zone : sert seulement à traiter d'abord les logements proches,
# tant que leur code postal n'est pas encore connu (la décision d'alerte se fait sur le code postal).
NEAR = (48.79, 48.92, 2.20, 2.50)

TRACKED = [("loyer_cc", "Loyer cc", "money"), ("categorie", "Catégorie", "text")]
CARD_KEYS = ("url", "resume", "type", "notes", "gamme", "ville", "code_postal", "photo")
CATEGORY_FR = {"social": "Logement social", "intermediaire": "Logement intermédiaire",
               "libre": "Logement à loyer libre", "etudiant": "Logement étudiant", "stationnement": "Stationnement"}


class CdcError(Exception):
    pass


# -----------------------------------------------------------------------------------------
# Accès au site (direct, ou via un navigateur si le site refuse les robots)
# -----------------------------------------------------------------------------------------

class Client:
    def __init__(self) -> None:
        self.mode = "direct"
        self.calls = 0
        self._opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        self._pw = self._browser = self._ctx = self._page = None

    def _direct(self, url: str, data: dict | None) -> tuple[int, str]:
        headers = {"User-Agent": config.USER_AGENT, "Accept-Language": "fr-FR,fr;q=0.9"}
        body = None
        if data is None:
            headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        else:  # même requête que la carte du site
            body = urllib.parse.urlencode(data).encode()
            headers.update({"Accept": "*/*", "X-Requested-With": "XMLHttpRequest", "Origin": BASE,
                            "Referer": BASE + LIST_PATH,
                            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"})
        req = urllib.request.Request(url, data=body, headers=headers)
        try:
            with self._opener.open(req, timeout=40) as resp:
                return resp.status, resp.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as err:
            return err.code, err.read().decode("utf-8", "replace")

    def _via_browser(self, url: str, data: dict | None) -> tuple[int, str]:
        if data is None:
            resp = self._page.goto(url, wait_until="domcontentloaded", timeout=60000)
            return (resp.status if resp else 0), self._page.content()
        resp = self._ctx.request.post(url, form=data, timeout=30000,
                                      headers={"X-Requested-With": "XMLHttpRequest", "Referer": BASE + LIST_PATH})
        return resp.status, resp.text()

    def start_browser(self) -> bool:
        try:
            from playwright.sync_api import sync_playwright
            self._pw = sync_playwright().start()
            self._browser = self._pw.chromium.launch(headless=True)
            self._ctx = self._browser.new_context(locale="fr-FR", user_agent=config.USER_AGENT)
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

    def fetch(self, url: str, data: dict | None = None, attempts: int = 4) -> str:
        last = "?"
        for attempt in range(attempts):
            if attempt:
                time.sleep(min(30, 2 ** attempt))
            self.calls += 1
            try:
                status, text = self._via_browser(url, data) if self.mode == "browser" else self._direct(url, data)
            except Exception as exc:
                last = f"{type(exc).__name__}: {exc}"
                continue
            if status == 200:
                return text
            last = f"HTTP {status}"
            if status == 429 or status >= 500:
                continue
            break
        raise CdcError(f"{url} : {last}")

    def get_bytes(self, url: str, max_bytes: int = 6_000_000) -> tuple[bytes, str]:
        if self.mode == "browser":
            resp = self._ctx.request.get(url, timeout=30000)
            if not resp.ok:
                raise CdcError(f"HTTP {resp.status}")
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

def num(value: Any) -> float | None:
    """« 1 013,62 € / mois » → 1013.62 ; « 1 013.62 EURO » → 1013.62."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    m = re.search(r"-?\d+(?:[.,]\d+)?", re.sub(r"\s", "", str(value)))
    return float(m.group(0).replace(",", ".")) if m else None


def squash(text: str | None) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def text_of(el) -> str | None:
    return (squash(el.get_text(" ")) or None) if el is not None else None


def parse_list(html: str) -> tuple[int | None, list[dict]]:
    """Renvoie (nombre de logements annoncé, liste des logements du bloc de données de la carte)."""
    lots: list[dict] = []
    total = None
    m = re.search(r"searchBootstrap\s*=\s*(\{)", html)
    if m:
        try:
            data, _ = json.JSONDecoder(strict=False).raw_decode(html, m.start(1))
        except json.JSONDecodeError:
            data = {}
        lots = [x for x in (data.get("lots") or []) if isinstance(x, dict)]
        total = data.get("lotCount") if isinstance(data.get("lotCount"), int) else None
    shown = re.search(r">\s*(\d[\d\s.]*?)\s+logements?\s+disponibles?", html)
    if shown:
        total = int(re.sub(r"\D", "", shown.group(1)))
    elif re.search(r"Aucun logement", html, re.I) and not lots:
        total = 0
    return (total if total is not None else (len(lots) or None)), lots


def lot_offer(lot: dict) -> dict:
    lid = str(lot.get("id_lot") or "").strip()
    offer = {
        "guid": lid,
        "id_lot": lid,
        "id_article": str(lot.get("id_article") or "").strip(),
        "categorie": str(lot.get("categorie_logement") or "").strip().lower(),
        "loyer_cc": num(lot.get("nb_prix")),
        "latitude": num(lot.get("latitude")),
        "longitude": num(lot.get("longitude")),
        "programme": str(lot.get("id_programme") or "").strip(),
    }
    return {k: v for k, v in offer.items() if v not in (None, "")}


def parse_card(fragment: str, lot_id: str) -> dict | None:
    """Vignette d'un logement (celle qui s'ouvre sur la carte du site)."""
    soup = BeautifulSoup(fragment or "", "html.parser")
    links = [a["href"] for a in soup.select("a[href]")
             if "/annonces-immobilieres/" in a["href"] or "/programme/" in a["href"]]
    url = next((x for x in links if x.rstrip("/").endswith("/" + lot_id)), links[0] if links else None)
    if url and url.startswith("/"):
        url = BASE + url
    card: dict[str, Any] = {"url": url}
    desc = soup.select_one(".description") or soup
    card["resume"] = text_of(desc.find(["h3", "h2", "h4"]))
    card["type"] = text_of(desc.select_one(".type"))
    card["notes"] = text_of(desc.select_one(".notes"))
    card["gamme"] = text_of(soup.select_one(".label"))
    where = text_of(desc.select_one(".location")) or ""
    m = re.match(r"(.*?)\s*\((\d{5})\)\s*$", where)
    card["ville"], card["code_postal"] = (m.group(1), m.group(2)) if m else (where or None, None)
    img = soup.select_one("img[src]")
    card["photo"] = img["src"] if img else None
    card = {k: v for k, v in card.items() if v not in (None, "")}
    return card if (card.get("url") or card.get("code_postal")) else None


def parse_detail(html: str) -> dict | None:
    soup = BeautifulSoup(html, "html.parser")
    info = soup.select_one("section.residenceInfo")
    if not info:  # logement retiré : le site renvoie vers l'accueil
        return None
    for other in soup.select(".similarHousings"):   # « autres logements similaires » : pas ce logement-ci
        other.decompose()
    d: dict[str, Any] = {}
    d["rubrique"] = text_of(info.find("h2"))
    d["titre"] = text_of(info.find("h1"))
    for p in info.find_all("p"):
        t = text_of(p) or ""
        if re.match(r"r[ée]f\b", t, re.I):
            d["reference"] = re.sub(r"^r[ée]f[^:]*:\s*", "", t, flags=re.I)
            break
    d["adresse_complete"] = text_of(info.find("address"))
    price = info.select_one("h3") or info.select_one(".h2")
    d["loyer_cc_affiche"] = num(text_of(price)) if price else None
    d["avantages"] = [t for t in (text_of(x) for x in info.select("strong")) if t]

    lines = []
    modal = soup.select_one("#aProposPrixModal")
    for row in (modal.select(".flex") if modal else []):
        cells = row.find_all("div", recursive=False)
        if not cells:
            continue
        label = text_of(cells[0]) or ""
        value = text_of(cells[1]) if len(cells) > 1 else None
        if label and not (label.lower() == "dont" and not value):
            lines.append({"libelle": label, "valeur": value})
    d["lignes_prix"] = lines

    d["gamme"] = text_of(soup.select_one(".residencePresentation .label"))
    plus = []
    for span in soup.select(".lesPlus .lesPlusDetails span"):
        t = text_of(span)
        if t and t not in plus:
            plus.append(t)
    d["les_plus"] = plus

    sections: dict[str, list[str]] = {}
    features = soup.select_one("section.features")
    for col in (features.select(".row > div") if features else []):
        head = text_of(col.find(["h3", "h4"]))
        items = [text_of(li) for ul in col.find_all("ul") if "reports" not in (ul.get("class") or [])
                 for li in ul.find_all("li")]
        if head:
            sections[head] = [i for i in items if i]
    d["caracteristiques"] = sections
    for li in soup.select("ul.reports li"):
        img = li.find("img")
        m = re.match(r"\s*(DPE|GES)\s*([A-G])?\b", (img.get("alt") or "") if img else "", re.I)
        if m:
            d[m.group(1).lower()] = {"classe": (m.group(2) or "").upper() or None, "valeur": num(text_of(li.find("b")))}

    desc = soup.select_one("section#description") or soup.select_one("section.description")
    if desc:
        for head in desc.find_all(["h2", "h3"]):
            head.decompose()
        lines_ = [squash(x) for x in desc.get_text("\n").split("\n")]
        d["description"] = "\n".join(x for x in lines_ if x) or None

    for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw = script.string or script.get_text() or ""
        try:
            ld = json.loads(raw, strict=False)
        except json.JSONDecodeError:
            ld = {}
            for key, pat in (("postalCode", r'"postalCode"\s*:\s*"(\d{5})"'), ("telephone", r'"telephone"\s*:\s*"([^"]+)"')):
                mm = re.search(pat, raw)
                if mm:
                    ld[key] = mm.group(1)
        if not isinstance(ld, dict):
            continue
        addr = ld.get("Address") or ld.get("address") or {}
        if isinstance(addr, dict):
            d["adresse"] = squash(addr.get("streetAddress")) or None
            d["ville"] = squash(addr.get("addressLocality")) or None
            d["code_postal"] = squash(addr.get("postalCode")) or None
        d["code_postal"] = d.get("code_postal") or ld.get("postalCode")
        d["telephone"] = ld.get("telephone")
        if not d.get("description") and squash(ld.get("description")):
            d["description"] = squash(ld.get("description"))
        break
    if not d.get("code_postal"):
        m = re.search(r"\((\d{5})\)", d.get("adresse_complete") or text_of(soup.select_one("section.location")) or "")
        d["code_postal"] = m.group(1) if m else None

    photos = []
    for img in soup.select(".residencePresentation img[src]"):
        src = img["src"]
        if "/photos/" in src and src not in photos:
            photos.append(src)
    d["photos"] = photos
    widget = soup.select_one(".alentoor-widget[data-config]")
    if widget:
        for key in ("latitude", "longitude"):
            mm = re.search(rf'"{key}"\s*:\s*(-?[\d.]+)', widget["data-config"])
            if mm:
                d[key] = float(mm.group(1))
    return {k: v for k, v in d.items() if v not in (None, "")}


def fetch_list(client: Client, expected_min: int) -> tuple[int, list[dict]]:
    url = BASE + LIST_PATH
    try:
        total, lots = parse_list(client.fetch(url))
        problem = None if len(lots) >= expected_min else f"{len(lots)} logement(s) dans la page"
    except CdcError as exc:
        total, lots, problem = None, [], str(exc)
    if problem and client.mode == "direct":
        log(f"⚠️  {SITE} : {problem}, bien moins que prévu (≥ {expected_min}) : "
            "le site refuse peut-être les robots.")
        if client.start_browser():
            total, lots = parse_list(client.fetch(url))
    if len(lots) < expected_min:
        raise CdcError(f"page incomplète : {len(lots)} logement(s) dans la page, au moins {expected_min} attendus")
    return (total or len(lots)), lots


def fetch_card(client: Client, lot_id: str, article_id: str | None) -> tuple[dict | None, str]:
    try:
        raw = client.fetch(BASE + CARD_PATH, {"idLot": lot_id, "idArticle": article_id or ""}, attempts=3)
        data = json.loads(raw)
    except (CdcError, ValueError) as exc:
        return None, f"vignette illisible : {exc}"
    card = parse_card(data.get("content") or "", lot_id) if isinstance(data, dict) else None
    return (card, "ok") if card else (None, "vignette vide")


def fetch_detail(client: Client, url: str) -> tuple[dict | None, str]:
    try:
        data = parse_detail(client.fetch(url, attempts=3))
    except CdcError as exc:
        return None, f"erreur : {exc}"
    return (data, "ok") if data else (None, "fiche introuvable")


def complete_record(client: Client, rec: dict, now: str) -> tuple[bool, str]:
    """Lit la vignette (si besoin) puis la fiche détaillée d'un logement."""
    liste = rec.setdefault("liste", {})
    if not liste.get("url"):
        time.sleep(config.REQUEST_DELAY)
        card, status = fetch_card(client, rec["guid"], liste.get("id_article"))
        if not card:
            return False, status
        liste.update(card)
    time.sleep(config.REQUEST_DELAY)
    data, status = fetch_detail(client, liste["url"])
    if not data:
        return False, status
    rec.update(fiche=data, fiche_statut="ok", fiche_recuperee_le=now, fiche_essais=0)
    return True, "ok"


# -----------------------------------------------------------------------------------------
# Champs (Excel + e-mail)
# -----------------------------------------------------------------------------------------

def L(rec: dict, key: str) -> Any:
    return (rec.get("liste") or {}).get(key)


def Fi(rec: dict, key: str) -> Any:
    return (rec.get("fiche") or {}).get(key)


def known_cp(rec: dict) -> str | None:
    return L(rec, "code_postal") or Fi(rec, "code_postal")


def in_zone_cp(cp: str | None) -> bool:
    cp = (cp or "").strip()
    return bool(cp) and any((cp.startswith(z[:-1]) if z.endswith("*") else cp == z) for z in ZONES)


def near(d: dict) -> bool:
    lat, lng = d.get("latitude"), d.get("longitude")
    return isinstance(lat, float) and isinstance(lng, float) and NEAR[0] <= lat <= NEAR[1] and NEAR[2] <= lng <= NEAR[3]


def no_alert_category(d: dict) -> bool:
    return (d.get("categorie") or "") in NO_ALERT


def rent_ok(rent: float | None) -> bool:
    return MAX_RENT is None or rent is None or rent <= MAX_RENT


def is_in_zone(rec: dict) -> bool:
    """Logement (pas un parking) dans une ville suivie, code postal connu : colonne et onglet « Zone d'alerte »."""
    return in_zone_cp(known_cp(rec)) and not no_alert_category(rec.get("liste") or {})


def offer_in_zone(offer: dict) -> bool:
    """Logement (pas un parking) dans la zone ; tant que la ville est inconnue, on se fie au GPS."""
    if no_alert_category(offer):
        return False
    return in_zone_cp(offer["code_postal"]) if offer.get("code_postal") else near(offer)


def angouleme(d: dict) -> bool:
    """Profil GrandAngoulême (autre destinataire) ; d = offre de la liste (vignette lue)."""
    return not no_alert_category(d) and profils.by_place(d.get("code_postal"), d.get("ville"))


def wants_alert(offer: dict) -> bool:
    if angouleme(offer):
        return profils.rent_ok(offer.get("loyer_cc"))
    return offer_in_zone(offer) and rent_ok(offer.get("loyer_cc"))


def alert_confirmed(rec: dict) -> bool | None:
    """Décision finale, sur le code postal : None = ville encore inconnue."""
    cp = known_cp(rec)
    if not cp:
        return None
    liste = rec.get("liste") or {}
    if profils.by_place(cp, liste.get("ville") or Fi(rec, "ville")):
        return not no_alert_category(liste) and profils.rent_ok(rent_cc(rec))
    return not no_alert_category(rec.get("liste") or {}) and in_zone_cp(cp) and rent_ok(rent_cc(rec))


def priority(rec: dict) -> bool:
    liste = rec.get("liste") or {}
    if no_alert_category(liste):
        return False
    return is_in_zone(rec) if known_cp(rec) else near(liste)


SMALL_WORDS = {"sur", "sous", "les", "le", "la", "de", "du", "des", "en", "et", "lès", "lez", "aux", "au", "d", "l"}


def nice_city(name: str | None, cp: str | None) -> str:
    name = squash(name)
    if not name:
        return ""
    if cp and cp.startswith("75") and name.upper().startswith("PARIS"):
        arr = int(cp[-2:]) if cp[-2:].isdigit() else 0
        return f"Paris {arr}{'er' if arr == 1 else 'e'}" if 1 <= arr <= 20 else "Paris"
    if name.isupper():
        words = name.lower().split()
        name = " ".join(w if (i and w in SMALL_WORDS) else w[:1].upper() + w[1:] for i, w in enumerate(words))
    return name


def city_label(rec: dict) -> str:
    return nice_city(L(rec, "ville") or Fi(rec, "ville"), known_cp(rec))


def rent_cc(rec: dict) -> float | None:
    return L(rec, "loyer_cc") or Fi(rec, "loyer_cc_affiche")


def resume_of(rec: dict) -> str | None:
    return Fi(rec, "titre") or L(rec, "resume")


def _parts(rec: dict) -> list[str]:
    return [p.strip() for p in re.split(r"\s+-\s+|\s*-\s*$", resume_of(rec) or "") if p.strip()]


def rooms(rec: dict) -> int | None:
    for p in _parts(rec):
        m = re.match(r"(\d+)\s*pi[eè]ce", p, re.I)
        if m:
            return int(m.group(1))
        if re.match(r"studio", p, re.I):
            return 1
    return None


def area(rec: dict) -> float | None:
    for p in _parts(rec):
        m = re.search(r"([\d.,]+)\s*m(²|2)", p)
        if m:
            return num(m.group(1))
    return None


def floor(rec: dict) -> str | None:
    for p in _parts(rec):
        if re.search(r"[ée]tage|rdc|rez", p, re.I):
            return p
    return None


def gamme(rec: dict) -> str | None:
    return Fi(rec, "gamme") or L(rec, "gamme") or CATEGORY_FR.get(L(rec, "categorie") or "", L(rec, "categorie"))


def kind(rec: dict) -> str | None:
    if L(rec, "type"):
        return L(rec, "type")
    rub = Fi(rec, "rubrique") or ""
    m = re.match(r"location\s+(.+)", rub, re.I)
    return m.group(1).capitalize() if m else None


def section_items(rec: dict, pattern: str) -> list[str]:
    for head, items in (Fi(rec, "caracteristiques") or {}).items():
        if re.search(pattern, head, re.I):
            return items or []
    return []


def equipment(rec: dict) -> str | None:
    items = []
    for x in section_items(rec, r"^logement") + list(Fi(rec, "les_plus") or []):
        if x not in items:
            items.append(x)
    return " ; ".join(items) or (L(rec, "notes") if not rec.get("fiche") else None)


def residence(rec: dict) -> str | None:
    items = [x for x in section_items(rec, r"r[ée]sidence") if not re.match(r"ann[ée]e de construction", x, re.I)]
    return " ; ".join(items) or None


def built(rec: dict) -> int | None:
    for x in section_items(rec, r"r[ée]sidence"):
        m = re.match(r"ann[ée]e de construction\s*:?\s*(\d{4})", x, re.I)
        if m:
            return int(m.group(1))
    return None


def heating(rec: dict) -> str | None:
    return " ; ".join(section_items(rec, r"chauffage")) or None


def other_sections(rec: dict) -> str | None:
    out = [f"{k} : {', '.join(v)}" for k, v in (Fi(rec, "caracteristiques") or {}).items()
           if v and not re.search(r"^logement|r[ée]sidence|chauffage", k, re.I)]
    return " ; ".join(out) or None


KNOWN_LINES = {
    "loyer_cc": r"^loyer charges comprises",
    "charges": r"^charges g[ée]n[ée]rales",
    "chauffage": r"^charges de chauffage",
    "depot": r"^d[ée]p[ôo]t de garantie",
    "energie": r"^facture d.[ée]nergie",
    "encadrement": r"encadrement des loyers",
    "ref_majore": r"loyer de r[ée]f[ée]rence major[ée]",
    "base": r"loyer de base",
    "complement": r"compl[ée]ment de loyer",
}


def price_line(rec: dict, key: str) -> dict | None:
    for line in Fi(rec, "lignes_prix") or []:
        if re.search(KNOWN_LINES[key], line.get("libelle") or "", re.I):
            return line
    return None


def price_value(rec: dict, key: str) -> float | None:
    line = price_line(rec, key)
    return num(line.get("valeur")) if line else None


def other_lines(rec: dict) -> str | None:
    out = []
    for line in Fi(rec, "lignes_prix") or []:
        lab = line.get("libelle") or ""
        if any(re.search(p, lab, re.I) for p in KNOWN_LINES.values()):
            continue
        out.append(f"{lab} : {line['valeur']}" if line.get("valeur") else lab)
    return " ; ".join(out) or None


def whole(value: Any) -> Any:
    return int(value) if isinstance(value, float) and value.is_integer() else value


def describe(offer: dict) -> str:
    what = offer.get("resume") or CATEGORY_FR.get(offer.get("categorie") or "", "logement")
    where = nice_city(offer.get("ville"), offer.get("code_postal")) or f"lot {offer.get('guid')}"
    return f"{what} · {where}"


def summary(rec: dict) -> dict:
    return {"reference": Fi(rec, "reference") or f"lot {rec.get('guid', '')}", "ville": city_label(rec),
            "code_postal": known_cp(rec) or "", "typologie": resume_of(rec) or gamme(rec) or "",
            "loyer": fmt(rent_cc(rec), "money") or ""}


STATUS_FR = {"en_ligne": "En ligne", "retiree": "Retirée"}


def F(section_, label, get, kind_="text", width=14, in_email=True):
    return Field(section_, label, get, kind_, (), width, in_email)


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
    F("Localisation", "Ville", city_label, width=24),
    F("Localisation", "Code postal", known_cp, width=10),
    F("Localisation", "Département", lambda r: (known_cp(r) or "")[:2] or None, width=8, in_email=False),
    F("Localisation", "Adresse", lambda r: Fi(r, "adresse"), width=30),
    F("Localisation", "Latitude", lambda r: L(r, "latitude") or Fi(r, "latitude"), width=10, in_email=False),
    F("Localisation", "Longitude", lambda r: L(r, "longitude") or Fi(r, "longitude"), width=10, in_email=False),
    F("Loyer et frais", "Loyer charges comprises (€/mois)", rent_cc, "money", 13),
    F("Loyer et frais", "Charges générales (€/mois)", lambda r: price_value(r, "charges"), "money", 12),
    F("Loyer et frais", "Charges de chauffage (€/mois)", lambda r: price_value(r, "chauffage"), "money", 12),
    F("Loyer et frais", "Dépôt de garantie (€)", lambda r: price_value(r, "depot"), "money", 12),
    F("Loyer et frais", "Facture d'énergie estimée (€/an)", lambda r: price_value(r, "energie"), "money", 12),
    F("Loyer et frais", "Encadrement des loyers", lambda r: bool(price_line(r, "encadrement")) if r.get("fiche") else None,
      "bool", 11),
    F("Loyer et frais", "Loyer de référence majoré (€/mois)", lambda r: price_value(r, "ref_majore"), "money", 13),
    F("Loyer et frais", "Loyer de base (€/mois)", lambda r: price_value(r, "base"), "money", 12),
    F("Loyer et frais", "Complément de loyer (€/mois)", lambda r: price_value(r, "complement"), "money", 12),
    F("Loyer et frais", "Honoraires", lambda r: " ; ".join(Fi(r, "avantages") or []) or None, width=14),
    F("Loyer et frais", "Autres montants", other_lines, width=24),
    F("Logement", "Catégorie", gamme, width=20),
    F("Logement", "Type", kind, width=12),
    F("Logement", "Pièces", rooms, "int", 7),
    F("Logement", "Surface (m²)", area, "area", 9),
    F("Logement", "Étage", floor, width=10),
    F("Logement", "Équipements", equipment, width=28),
    F("Logement", "Résidence", residence, width=22),
    F("Logement", "Année de construction", built, "int", 10),
    F("Logement", "Chauffage", heating, width=18),
    F("Logement", "Autres caractéristiques", other_sections, width=22),
    F("Énergie", "DPE (classe)", lambda r: (Fi(r, "dpe") or {}).get("classe"), width=8),
    F("Énergie", "Consommation (kWh/m²/an)", lambda r: whole((Fi(r, "dpe") or {}).get("valeur")), "int", 11),
    F("Énergie", "GES (classe)", lambda r: (Fi(r, "ges") or {}).get("classe"), width=8),
    F("Énergie", "Émissions (kg CO₂/m²/an)", lambda r: whole((Fi(r, "ges") or {}).get("valeur")), "int", 11),
    F("Description", "Titre", resume_of, width=28),
    F("Description", "Description", lambda r: Fi(r, "description"), "longtext", 60),
    F("Description", "Photos", lambda r: len(Fi(r, "photos") or []) or (1 if L(r, "photo") else 0), "int", 7, False),
    F("Références", "Lien vers l'annonce", lambda r: L(r, "url"), "link", 14, False),
    F("Références", "Référence", lambda r: Fi(r, "reference"), width=14),
    F("Références", "N° de lot", lambda r: r.get("guid"), width=10),
    F("Références", "Téléphone de l'agence", lambda r: Fi(r, "telephone"), width=13),
    F("Références", "Programme (identifiant)", lambda r: L(r, "programme"), width=10, in_email=False),
]


def profile(store: Store):
    import export_excel
    urls = {g: L(r, "url") for g, r in store.records.items()}
    return export_excel.Profile(
        title="Suivi des offres — cdc-habitat.fr (location)", fields=FIELDS, extra_fn=lambda r: {}, zone_fn=is_in_zone,
        detail_url=lambda guid: urls.get(guid) or PUBLIC_LIST,
        notes=[("Zone d'alerte (codes postaux)", ", ".join(ZONES)),
               ("Plafond des alertes", f"{fmt(MAX_RENT, 'money')} par mois charges comprises" if MAX_RENT else "aucun"),
               ("Jamais d'alerte pour", ", ".join(CATEGORY_FR.get(c, c) for c in sorted(NO_ALERT)) or "—"),
               ("Fiches en cours de lecture", "Ville, adresse et détails se remplissent au fil des passages "
                                              f"({MAX_DETAILS_PER_RUN} logements par passage)."),
               ("Source", f"Page de résultats {BASE}{LIST_PATH} (liste complète), vignettes et fiches détaillées")])


# -----------------------------------------------------------------------------------------
# E-mail
# -----------------------------------------------------------------------------------------

ACCENT = "#00609c"


def render(rec: dict, cids: list[str], relisted: bool = False) -> tuple[str, str, str]:
    from alertes import INK, LINE, MUTED, h

    rent = rent_cc(rec)
    city = city_label(rec) or "ville inconnue"
    n = rooms(rec)
    bits = [f"{n} pièce{'s' if n and n > 1 else ''}" if n else None, fmt(area(rec), "area"),
            f"{fmt(rent, 'money')}/mois cc" if rent else None]
    prefix = "🔁 CDC Habitat · de nouveau en ligne" if relisted else "🏠 CDC Habitat · nouvelle annonce"
    subj = f"{prefix} · {city} · " + " · ".join(b for b in bits if b)

    detected = parse_dt(rec.get("premiere_detection"))
    when = f"détectée le {(detected or parse_dt(now_iso())).strftime('%d/%m/%Y à %H:%M')}"
    kicker = "CDC Habitat · De nouveau en ligne" if relisted else "CDC Habitat · Nouvelle annonce"
    head_title = f"{kind(rec) or 'Logement'} · {resume_of(rec)}" if resume_of(rec) else f"Logement à {city}"
    address = Fi(rec, "adresse_complete") or f"{city} ({known_cp(rec) or '?'})"
    url = L(rec, "url") or PUBLIC_LIST
    by_label = {f.label: fmt(f.get(rec), f.kind) for f in FIELDS}
    facts = [("Pièces", by_label.get("Pièces")), ("Surface", by_label.get("Surface (m²)")),
             ("Étage", by_label.get("Étage"))]
    facts = [(k, v) for k, v in facts if v]
    groups = [
        ("Loyer et frais", ["Charges générales (€/mois)", "Charges de chauffage (€/mois)", "Dépôt de garantie (€)",
                            "Facture d'énergie estimée (€/an)", "Encadrement des loyers",
                            "Loyer de référence majoré (€/mois)", "Loyer de base (€/mois)",
                            "Complément de loyer (€/mois)", "Honoraires", "Autres montants"]),
        ("Le logement", ["Catégorie", "Type", "Équipements", "Résidence", "Année de construction", "Chauffage",
                         "Autres caractéristiques"]),
        ("Énergie", ["DPE (classe)", "Consommation (kWh/m²/an)", "GES (classe)", "Émissions (kg CO₂/m²/an)"]),
        ("Références", ["Référence", "N° de lot", "Téléphone de l'agence"]),
    ]
    short = lambda k: re.sub(r" \((€/mois|€)\)$", "", k)
    sections = [(t, [(short(k), by_label[k]) for k in keys if by_label.get(k)]) for t, keys in groups]
    sections = [(t, rows) for t, rows in sections if rows]
    description = Fi(rec, "description")

    lines = [f"{kicker} — {when}", "", head_title, address]
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
        f'<td style="width:6px"></td>' for k, v in facts)

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
  <tr><td style="padding:0 0 14px 0;font-size:14px;color:{MUTED}">{h(address)}</td></tr>
  {photos_html}
  {rent_html}
  <tr><td style="padding:12px 0 0 0"><table role="presentation" cellpadding="0" cellspacing="0"><tr>{facts_html}</tr></table></td></tr>
  <tr><td style="padding:20px 0 0 0"><a href="{h(url)}" style="display:inline-block;padding:12px 20px;background:{ACCENT};
    color:#ffffff;text-decoration:none;border-radius:8px;font-size:15px;font-weight:700">Voir l'annonce sur cdc-habitat.fr</a></td></tr>
  {no_detail}
  {top}
  {desc_html}
  {bottom}
  <tr><td style="padding:26px 0 0 0;font-size:12px;color:{MUTED}">
    Alerte envoyée par votre veille CDC Habitat · <a href="{h(url)}" style="color:{MUTED}">{h(url)}</a></td></tr>
</table></td></tr></table></td></tr></table></body></html>"""
    return subj, text, html


def photo_urls(rec: dict) -> list[str]:
    return list(Fi(rec, "photos") or []) or ([L(rec, "photo")] if L(rec, "photo") else [])


def send_alert(client: Client, rec: dict, relisted: bool = False) -> None:
    import alertes
    photos = alertes.download_photos(client, photo_urls(rec))
    liste = rec.get("liste") or {}
    to = profils.TO if profils.by_place(known_cp(rec), liste.get("ville") or Fi(rec, "ville")) else None
    alertes.send_rendered(lambda cids: render(rec, cids, relisted), photos, sender_name=SENDER, to=to)


# -----------------------------------------------------------------------------------------
# Passage
# -----------------------------------------------------------------------------------------

def open_store() -> Store:
    return Store(DATA_DIR, summary=summary, tracked=TRACKED)


def export(store: Store) -> None:
    import export_excel
    export_excel.build(store, EXPORT_PATH, profile(store))
    log(f"Excel CDC Habitat mis à jour : {EXPORT_PATH}")


def run() -> int:
    if not ENABLED:
        return 0
    store = open_store()
    client = Client()
    exit_code = 0
    changed = False
    try:
        now = now_iso()
        last_total = store.state.get("dernier_total")
        expected = max(20, int(0.5 * last_total)) if last_total else MIN_EXPECTED
        try:
            total, lots = fetch_list(client, expected)
        except CdcError as exc:
            return handle_failure(store, str(exc), SITE, PUBLIC_LIST)
        offers: dict[str, dict] = {}
        for lot in lots:
            offer = lot_offer(lot)
            if offer.get("guid"):
                offers[offer["guid"]] = offer
        complete = total > 0 and len(offers) >= config.MIN_COMPLETENESS * total
        log(f"{SITE} : {len(offers)} logements dans la liste, {total} annoncés (mode {client.mode}).")

        # ce que l'on sait déjà de chaque logement (ville, lien…) est conservé
        for guid, offer in offers.items():
            rec = store.records.get(guid)
            if rec:
                for key in CARD_KEYS:
                    if key in (rec.get("liste") or {}):
                        offer.setdefault(key, rec["liste"][key])

        # nouveautés : on lit leur vignette tout de suite pour connaître la ville avant de décider d'une alerte
        if store.state.get("initialise"):
            new = [o for g, o in offers.items() if g not in store.records]
            new.sort(key=lambda o: not (near(o) and not no_alert_category(o)))
            for offer in new[:MAX_NEW_CARDS]:
                time.sleep(config.REQUEST_DELAY)
                card, status = fetch_card(client, offer["guid"], offer.get("id_article"))
                if card:
                    offer.update(card)
                else:
                    log(f"⚠️  {SITE} : lot {offer['guid']} — {status} (ville vérifiée plus tard).")

        counts = reconcile(store, offers, now, complete=complete, total=total,
                           in_zone=offer_in_zone, wants_alert=wants_alert, describe=describe,
                           rent_of=lambda o: o.get("loyer_cc"), max_rent=MAX_RENT)
        if complete:
            store.set_state(dernier_total=total)

        # vignettes et fiches détaillées (les logements proches de la zone d'abord)
        queue, budget = detail_queue(store, priority, MAX_DETAILS_PER_RUN)
        ok = failures = 0
        for rec in queue[:budget]:
            done, status = complete_record(client, rec, now)
            failures = 0 if done else failures + 1
            if done:
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

        # alertes (la ville est revérifiée sur le code postal avant l'envoi)
        pending = pending_alerts(store)
        sent = 0
        if pending and not config.EMAIL_ENABLED:
            log(f"⚠️  {len(pending)} alerte(s) {SITE} en attente mais l'e-mail n'est pas configuré.")
        elif pending:
            for rec in pending[: config.MAX_ALERTS_PER_RUN]:
                if not rec.get("fiche"):
                    complete_record(client, rec, now)
                    store.touch(rec["guid"])
                decision = alert_confirmed(rec)
                if decision is None:
                    tries = (rec.get("alerte_essais") or 0) + 1
                    rec["alerte_essais"] = tries
                    store.touch(rec["guid"])
                    if tries >= 5:
                        rec["alerte"] = "abandonnee"
                        log(f"⚠️  {SITE} : lot {rec['guid']} — ville introuvable, alerte abandonnée.")
                    else:
                        log(f"{SITE} : lot {rec['guid']} — ville pas encore connue, alerte reportée.")
                    continue
                if not decision:
                    rec["alerte"] = "hors_zone"
                    store.touch(rec["guid"])
                    log(f"{SITE} : {describe(rec['liste'])} — hors zone ou au-dessus du plafond, pas d'alerte.")
                    continue
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
        bilan(SITE, counts, sent, online, len(offers))
        if changed or not EXPORT_PATH.exists():
            export(store)
        log(f"{SITE} : {client.calls} requête(s).")
        return exit_code
    finally:
        client.close()


def run_test_email() -> int:
    """Exemple d'alerte avec une vraie annonce CDC Habitat de la zone (la moins chère trouvée)."""
    import alertes
    client = Client()
    try:
        _, lots = fetch_list(client, 1)
        offers = [lot_offer(lot) for lot in lots]
        nearby = sorted((o for o in offers if near(o) and not no_alert_category(o)),
                        key=lambda o: o.get("loyer_cc") or 1e9)
        chosen = None
        for offer in nearby[:15]:
            time.sleep(config.REQUEST_DELAY)
            card, _ = fetch_card(client, offer["guid"], offer.get("id_article"))
            if card and in_zone_cp(card.get("code_postal")):
                chosen = {**offer, **card}
                break
        if not chosen:
            alertes.send_simple("Test — Veille CDC Habitat", "Les e-mails fonctionnent 👍 (aucune annonce CDC Habitat "
                                "trouvée dans vos villes pour illustrer l'alerte).", PUBLIC_LIST, sender_name=SENDER)
            log("E-mail de test CDC Habitat simple envoyé.")
            return 0
        rec = {"guid": chosen["guid"], "premiere_detection": now_iso(), "liste": chosen, "fiche": None}
        complete_record(client, rec, now_iso())
        send_alert(client, rec)
        log(f"E-mail de test CDC Habitat envoyé avec l'annonce : {describe(chosen)}")
        return 0
    finally:
        client.close()


def run_export() -> int:
    export(open_store())
    return 0
