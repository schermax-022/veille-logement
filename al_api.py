"""Client de l'API publique du site logement-actionlogement.fr.

Deux appels, identifiés dans le trafic du site :
  POST /demands/public/offers-overview?size=…&page=…   liste paginée des offres
  POST /demands/public/offer-showcase-details           fiche détaillée {"offerGuid": …}

Par défaut, l'API est interrogée directement (rapide, sans navigateur). Si elle refuse les
requêtes qui ne viennent pas d'un navigateur (protection Cloudflare), le client bascule sur un
Chromium sans fenêtre qui fait les mêmes appels que le site (nécessite USE_BROWSER=true sur GitHub).
"""

from __future__ import annotations

import json
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field

import config
from config import API_BASE, SITE_BASE


def log(msg: str) -> None:
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


class ApiError(Exception):
    def __init__(self, message: str, status: int | None = None):
        super().__init__(message)
        self.status = status


class ApiBlocked(ApiError):
    pass


_FETCH_JS = """
async ({url, method, body}) => {
  const opts = {method, headers: {'Accept': 'application/json, text/plain, */*'}};
  if (body !== null) { opts.headers['Content-Type'] = 'application/json'; opts.body = body; }
  const r = await fetch(url, opts);
  return {status: r.status, text: await r.text(), ctype: r.headers.get('content-type') || ''};
}
"""


class ApiClient:
    def __init__(self) -> None:
        self.mode = "direct"
        self.calls = 0
        self._pw = self._browser = self._ctx = self._page = None
        if config.USE_BROWSER:
            self._start_browser()

    # -- transport -----------------------------------------------------------------
    def _headers(self, with_body: bool) -> dict:
        h = {
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "fr-FR,fr;q=0.9",
            "Origin": SITE_BASE,
            "Referer": SITE_BASE + "/",
            "User-Agent": config.USER_AGENT,
        }
        if with_body:
            h["Content-Type"] = "application/json"
        return h

    def _direct(self, method: str, url: str, body: str | None) -> tuple[int, str, str]:
        data = body.encode("utf-8") if body is not None else None
        req = urllib.request.Request(url, data=data, method=method, headers=self._headers(body is not None))
        try:
            with urllib.request.urlopen(req, timeout=40) as resp:
                return resp.status, resp.read().decode("utf-8", "replace"), resp.headers.get("Content-Type", "")
        except urllib.error.HTTPError as err:
            return err.code, err.read().decode("utf-8", "replace"), err.headers.get("Content-Type", "")

    def _start_browser(self) -> None:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise ApiBlocked("L'API refuse les requêtes directes et Playwright n'est pas installé : "
                             "sur GitHub, ajoutez la variable USE_BROWSER = true ; en local, "
                             "pip install -r requirements-navigateur.txt.") from exc
        try:
            self._pw = sync_playwright().start()
            self._browser = self._pw.chromium.launch(headless=True)
        except Exception as exc:
            raise ApiBlocked("Chromium n'est pas installé : lancez « python -m playwright install --with-deps "
                             "chromium » (sur GitHub : ajoutez la variable USE_BROWSER = true).") from exc
        self._ctx = self._browser.new_context(locale="fr-FR", user_agent=config.USER_AGENT)
        self._page = self._ctx.new_page()
        self._page.goto(SITE_BASE + "/search", wait_until="domcontentloaded", timeout=60000)
        self._page.wait_for_timeout(1500)
        self.mode = "browser"
        log("Mode navigateur activé (Chromium sans fenêtre).")

    def _browser_fetch(self, method: str, url: str, body: str | None) -> tuple[int, str, str]:
        res = self._page.evaluate(_FETCH_JS, {"url": url, "method": method, "body": body})
        return res["status"], res["text"], res["ctype"]

    def close(self) -> None:
        for obj in (self._browser, self._pw):
            try:
                if obj is self._pw and obj:
                    obj.stop()
                elif obj:
                    obj.close()
            except Exception:
                pass

    # -- appels ----------------------------------------------------------------------
    def request_json(self, method: str, path: str, body: dict | None = None, params: dict | None = None) -> dict:
        url = API_BASE + path + ("?" + urllib.parse.urlencode(params) if params else "")
        payload = json.dumps(body, ensure_ascii=False) if body is not None else None
        last = "?"
        for attempt in range(5):
            if attempt:
                time.sleep(min(30, 2 ** attempt))
            self.calls += 1
            try:
                if self.mode == "direct":
                    status, text, ctype = self._direct(method, url, payload)
                else:
                    status, text, ctype = self._browser_fetch(method, url, payload)
            except Exception as exc:  # réseau coupé, délai dépassé…
                last = f"{type(exc).__name__}: {exc}"
                continue

            is_json = text.lstrip()[:1] in ("{", "[")
            if status == 200 and is_json:
                return json.loads(text)
            if not is_json and status in (200, 401, 403, 503):  # page HTML de protection anti-robots
                last = f"HTTP {status} (accès refusé ou page de protection)"
                if self.mode == "direct":
                    log(f"⚠️  L'API refuse les requêtes directes ({last}) : bascule sur le navigateur.")
                    self._start_browser()
                    continue
                raise ApiBlocked(f"L'API refuse l'accès même via le navigateur : {last}", status)
            if status == 429 or status >= 500:
                last = f"HTTP {status}"
                continue
            raise ApiError(f"HTTP {status} sur {path} : {text[:300]}", status)
        raise ApiError(f"Échec après plusieurs essais sur {path} : {last}")

    def get_bytes(self, url: str, max_bytes: int = 6_000_000) -> tuple[bytes, str]:
        """Télécharge une photo (les liens des photos expirent au bout de quelques heures)."""
        if self.mode == "browser":
            resp = self._ctx.request.get(url, timeout=30000)
            if not resp.ok:
                raise ApiError(f"HTTP {resp.status}", resp.status)
            return resp.body()[:max_bytes], resp.headers.get("content-type", "")
        req = urllib.request.Request(url, headers={"User-Agent": config.USER_AGENT, "Referer": SITE_BASE + "/"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read(max_bytes + 1)[:max_bytes], resp.headers.get("Content-Type", "")


# ---------------------------------------------------------------------------------
# Récupération des offres
# ---------------------------------------------------------------------------------

@dataclass
class Crawl:
    offers: dict = field(default_factory=dict)   # guid -> offre (liste)
    total: int = 0                               # nombre annoncé par l'API
    pages: int = 0
    size: int = 0                                # taille de page acceptée par l'API
    errors: list = field(default_factory=list)

    @property
    def complete(self) -> bool:
        return (not self.errors and self.total > 0
                and len(self.offers) >= config.MIN_COMPLETENESS * self.total)


OVERVIEW = "/demands/public/offers-overview"
DETAILS = "/demands/public/offer-showcase-details"


def fetch_all_offers(client: ApiClient, size: int | None = None) -> Crawl:
    crawl = Crawl()
    size = size or config.PAGE_SIZE
    try:
        first = client.request_json("POST", OVERVIEW, config.SEARCH_BODY, {"size": size, "page": 0})
    except ApiError as exc:
        if size == 20 or isinstance(exc, ApiBlocked) or not exc.status or exc.status >= 500:
            raise
        log(f"Taille de page {size} refusée ({exc.status}) : on repasse à 20.")
        size = 20
        first = client.request_json("POST", OVERVIEW, config.SEARCH_BODY, {"size": size, "page": 0})

    crawl.size = size
    crawl.total = int(first.get("totalSize") or 0)
    crawl.pages = int(first.get("totalPages") or 1)

    def add(resp: dict) -> int:
        n = 0
        for offer in resp.get("data") or []:
            if isinstance(offer, dict) and offer.get("guid"):
                crawl.offers[offer["guid"]] = offer
                n += 1
        return n

    add(first)
    last_page = min(crawl.pages, config.MAX_PAGES)
    if crawl.pages > config.MAX_PAGES:
        crawl.errors.append(f"{crawl.pages} pages annoncées, limite MAX_PAGES = {config.MAX_PAGES}")
    for page in range(1, last_page):
        time.sleep(config.REQUEST_DELAY)
        try:
            resp = client.request_json("POST", OVERVIEW, config.SEARCH_BODY, {"size": size, "page": page})
        except ApiBlocked:
            raise
        except ApiError as exc:
            crawl.errors.append(f"page {page} : {exc}")
            continue
        if not add(resp) and page < last_page - 1:
            crawl.errors.append(f"page {page} vide")
    return crawl


def fetch_detail(client: ApiClient, guid: str) -> tuple[dict | None, str]:
    try:
        resp = client.request_json("POST", DETAILS, {"offerGuid": guid})
    except ApiBlocked:
        raise
    except ApiError as exc:
        return None, f"erreur : {exc}"
    data = resp.get("data")
    if isinstance(data, dict) and data:
        return data, "ok"
    return None, f"fiche vide ({resp.get('code')} {resp.get('message')})"
