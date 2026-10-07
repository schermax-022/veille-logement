#!/usr/bin/env python3
"""
Veille des offres de logement-actionlogement.fr et de inli.fr

À chaque passage, pour chaque site (le détail In'li est dans inli.py) :
  1. récupère toutes les offres publiées (API publique du site, toute la France par défaut) ;
  2. met à jour la base data/ : nouvelles offres, retraits, remises en ligne, modifications ;
  3. récupère les fiches détaillées (en priorité celles de la zone d'alerte) ;
  4. envoie un e-mail par nouvelle offre dans la zone d'alerte, reproduisant l'annonce ;
  5. régénère le fichier Excel si quelque chose a changé.

Usage :
    python veille.py               # passage normal
    python veille.py --test-email  # envoie un exemple d'alerte avec une vraie annonce parisienne
    python veille.py --export      # régénère seulement les fichiers Excel
    python veille.py --site inli   # un seul site (action-logement ou inli)
"""

from __future__ import annotations

import argparse
import sys
import time

import config
from al_api import ApiBlocked, ApiClient, ApiError, fetch_all_offers, fetch_detail, log
from annonce import fmt, is_in_zone, photo_urls, pick, title
from stockage import Store, clean, now_iso
from suivi import bilan, detail_queue, handle_failure, pending_alerts, reconcile


def zone_of_offer(offer: dict) -> bool:
    return config.in_alert_zone(((offer.get("lodging") or {}).get("municipality") or {}).get("code"))


# Plafond de loyer pour les alertes, en €/mois : c'est le loyer affiché par le site (montant total
# mensuel, celui que filtre l'option « loyer max » du site). Réglable avec la variable ALERT_MAX_RENT ;
# « 0 » = pas de plafond. Une offre sans loyer indiqué déclenche quand même l'alerte, par prudence.
_max_rent = config.env("ALERT_MAX_RENT", "1100").replace(",", ".").replace("€", "").strip()
ALERT_MAX_RENT = float(_max_rent) if _max_rent not in ("", "0") else None


def offer_rent(offer: dict) -> float | None:
    for key in ("totalRentAmountBaseBound", "housingFeeBaseBound", "totalRentAmountMax"):
        value = offer.get(key)
        if isinstance(value, (int, float)) and value > 0:
            return float(value)
    return None


def wants_alert(offer: dict) -> bool:
    """Alerte = dans une ville suivie ET loyer sous le plafond."""
    if not zone_of_offer(offer):
        return False
    rent = offer_rent(offer)
    return ALERT_MAX_RENT is None or rent is None or rent <= ALERT_MAX_RENT


# -----------------------------------------------------------------------------------------
# Passage normal
# -----------------------------------------------------------------------------------------

def run_check() -> int:
    store = Store(config.DATA_DIR)
    client = ApiClient()
    exit_code = 0
    changed = False
    site = "Action Logement"
    try:
        now = now_iso()
        try:
            crawl = fetch_all_offers(client, store.state.get("taille_page"))
        except (ApiError, ApiBlocked) as exc:
            return handle_failure(store, str(exc), site, config.SITE_BASE + "/search")
        log(f"{site} : {len(crawl.offers)} offres récupérées sur {crawl.total} annoncées "
            f"({crawl.pages} page(s), mode {client.mode}).")
        for err in crawl.errors:
            log(f"⚠️  {err}")
        if not crawl.offers:
            return handle_failure(store, "; ".join(crawl.errors) or "liste vide", site, config.SITE_BASE + "/search")

        # 1-2. nouvelles offres, modifications, remises en ligne, retraits
        counts = reconcile(store, crawl.offers, now, complete=crawl.complete, total=crawl.total,
                           in_zone=zone_of_offer, wants_alert=wants_alert,
                           describe=lambda o: title({"liste": o}), rent_of=offer_rent, max_rent=ALERT_MAX_RENT)
        store.set_state(taille_page=crawl.size)

        # 3. fiches détaillées
        fresh: dict[str, dict] = {}
        queue, budget = detail_queue(store, is_in_zone, config.MAX_DETAILS_PER_RUN)
        ok = failures_in_a_row = 0
        for rec in queue[:budget]:
            time.sleep(config.REQUEST_DELAY)
            try:
                data, status = fetch_detail(client, rec["guid"])
            except ApiBlocked as exc:
                log(f"⚠️  Lecture des fiches interrompue : {exc}")
                break
            except Exception as exc:  # une fiche illisible ne doit pas arrêter le passage
                data, status = None, f"erreur : {exc}"
            failures_in_a_row = 0 if data else failures_in_a_row + 1
            if failures_in_a_row >= 5:
                log("⚠️  Le site répond mal : lecture des fiches reportée au prochain passage.")
                break
            if data:
                fresh[rec["guid"]] = data
                rec.update(fiche=clean(data), fiche_statut="ok", fiche_recuperee_le=now, fiche_essais=0)
                ok += 1
            else:
                rec.update(fiche_statut=status, fiche_essais=(rec.get("fiche_essais") or 0) + 1)
            store.touch(rec["guid"])
        if queue:
            log(f"Fiches détaillées : {ok} récupérée(s), "
                f"{len(queue) - ok} en attente pour les prochains passages.")

        changed |= store.save()

        # 4. alertes
        pending = pending_alerts(store)
        sent = 0
        if pending and not config.EMAIL_ENABLED:
            log(f"⚠️  {len(pending)} alerte(s) en attente mais l'e-mail n'est pas configuré (SMTP_HOST, EMAIL_TO).")
        elif pending:
            import alertes
            for rec in pending[: config.MAX_ALERTS_PER_RUN]:
                raw = fresh.get(rec["guid"])
                if raw is None:  # les liens des photos expirent : on relit la fiche pour en avoir des frais
                    try:
                        raw, _ = fetch_detail(client, rec["guid"])
                    except Exception:
                        raw = None
                    if raw:
                        rec.update(fiche=clean(raw), fiche_statut="ok", fiche_recuperee_le=now)
                urls = photo_urls({"fiche": raw, "liste": crawl.offers.get(rec["guid"])})
                photos = alertes.download_photos(client, urls)
                relisted = rec["alerte"] == "a_envoyer_remise"
                try:
                    alertes.send_offer_alert(rec, photos, relisted)
                except Exception as exc:
                    log(f"⚠️  Alerte non envoyée ({title(rec)}) : {exc} — nouvel essai au prochain passage.")
                    exit_code = 1
                    break
                rec.update(alerte="envoyee", alerte_envoyee_le=now_iso())
                store.touch(rec["guid"])
                sent += 1
                log(f"✉️  Alerte envoyée : {title(rec)} — "
                    f"{fmt(pick(rec, 'totalRentAmountBaseBound'), 'money') or 'loyer non indiqué'}")
            changed |= store.save()

        # 5. bilan et export
        online = sum(1 for r in store.records.values() if r["statut"] == "en_ligne")
        bilan(site, counts, sent, online, len(crawl.offers))
        if changed or not config.EXPORT_PATH.exists():
            import export_excel
            export_excel.build(store, config.EXPORT_PATH)
            log(f"Excel mis à jour : {config.EXPORT_PATH}")
        log(f"{client.calls} appel(s) à l'API.")
        return exit_code
    finally:
        client.close()


# -----------------------------------------------------------------------------------------
# Autres modes
# -----------------------------------------------------------------------------------------

PARIS_BODY = {"municipalities": [{"code": f"751{n:02d}", "postcode": f"750{n:02d}"} for n in range(1, 21)],
              "searchRadiusInKm": 0, "typologyCodes": [], "productGuids": []}


def run_test_email() -> int:
    if not config.EMAIL_ENABLED:
        log("E-mail non configuré : renseignez SMTP_HOST, SMTP_USER, SMTP_PASSWORD, EMAIL_TO.")
        return 1
    import alertes
    client = ApiClient()
    try:
        resp = client.request_json("POST", "/demands/public/offers-overview", PARIS_BODY, {"size": 20, "page": 0})
        offers = resp.get("data") or []
        if not offers:
            alertes.send_simple("Test — Veille Action Logement", "Les e-mails fonctionnent 👍 (aucune offre "
                                "parisienne trouvée pour illustrer l'alerte).", config.SITE_BASE + "/search")
            log("E-mail de test simple envoyé.")
            return 0
        offer = offers[0]
        raw, _ = fetch_detail(client, offer["guid"])
        rec = {"guid": offer["guid"], "premiere_detection": now_iso(), "liste": clean(offer),
               "fiche": clean(raw) if raw else None}
        photos = alertes.download_photos(client, photo_urls({"fiche": raw, "liste": offer}))
        alertes.send_offer_alert(rec, photos)
        log(f"E-mail de test envoyé avec l'annonce : {title(rec)}")
        return 0
    finally:
        client.close()


def run_export() -> int:
    import export_excel
    store = Store(config.DATA_DIR)
    export_excel.build(store, config.EXPORT_PATH)
    log(f"Excel régénéré : {config.EXPORT_PATH} ({len(store.records)} offres)")
    return 0


def _safe(label: str, fn) -> int:
    """Lance un site ; une panne d'un site n'empêche pas l'autre de tourner."""
    import traceback
    try:
        return fn()
    except Exception:
        log(f"❌ {label} : erreur inattendue")
        traceback.print_exc()
        return 1


def main() -> int:
    import inli
    parser = argparse.ArgumentParser(description="Veille des offres Action Logement et In'li")
    parser.add_argument("--test-email", action="store_true", help="envoie un exemple d'alerte par site")
    parser.add_argument("--export", action="store_true", help="régénère seulement les fichiers Excel")
    parser.add_argument("--site", choices=["tous", "action-logement", "inli"], default="tous")
    args = parser.parse_args()
    if args.test_email:
        al, il = run_test_email, inli.run_test_email
    elif args.export:
        al, il = run_export, inli.run_export
    else:
        al, il = run_check, inli.run
    codes = []
    if args.site in ("tous", "action-logement"):
        codes.append(_safe("Action Logement", al))
    if args.site in ("tous", "inli") and inli.ENABLED:
        codes.append(_safe("In'li", il))
    return max(codes or [0])


if __name__ == "__main__":
    sys.exit(main())
