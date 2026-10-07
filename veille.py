#!/usr/bin/env python3
"""
Veille des offres de logement-actionlogement.fr

À chaque passage :
  1. récupère toutes les offres publiées (API publique du site, toute la France par défaut) ;
  2. met à jour la base data/ : nouvelles offres, retraits, remises en ligne, modifications ;
  3. récupère les fiches détaillées (en priorité celles de la zone d'alerte) ;
  4. envoie un e-mail par nouvelle offre dans la zone d'alerte, reproduisant l'annonce ;
  5. régénère le fichier Excel si quelque chose a changé.

Usage :
    python veille.py               # passage normal
    python veille.py --test-email  # envoie un exemple d'alerte avec une vraie annonce parisienne
    python veille.py --export      # régénère seulement le fichier Excel
"""

from __future__ import annotations

import argparse
import sys
import time

import config
from al_api import ApiBlocked, ApiClient, ApiError, fetch_all_offers, fetch_detail, log
from annonce import fmt, is_in_zone, parse_dt, photo_urls, pick, title
from stockage import Store, clean, now_iso


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


def fmt_dt(iso: str | None) -> str:
    dt = parse_dt(iso)
    return dt.strftime("%d/%m/%Y %H:%M") if dt else "—"


# -----------------------------------------------------------------------------------------
# Passage normal
# -----------------------------------------------------------------------------------------

def handle_failure(store: Store, reason: str) -> int:
    """Échec de récupération : on prévient une seule fois après plusieurs échecs d'affilée."""
    n = (store.state.get("passages_vides") or 0) + 1
    store.set_state(passages_vides=n)
    log(f"⚠️  Aucune offre récupérée ({reason}) — {n} passage(s) en échec d'affilée.")
    code = 0
    if n >= config.EMPTY_ALERT_AFTER and not store.state.get("alerte_vide_envoyee"):
        if config.EMAIL_ENABLED:
            import alertes
            try:
                alertes.send_simple(
                    "⚠️ Veille Action Logement : plus aucune offre récupérée",
                    f"Depuis {n} vérifications, le script ne récupère plus les offres.\n\nDernière erreur : {reason}\n\n"
                    "Si l'erreur parle d'accès refusé, ajoutez la variable USE_BROWSER = true sur GitHub.",
                    config.SITE_BASE + "/search")
            except Exception as exc:
                log(f"⚠️  E-mail d'avertissement non envoyé : {exc}")
        store.set_state(alerte_vide_envoyee=True)
        code = 1  # un seul échec visible côté GitHub, pas un toutes les 20 minutes
    store.save()
    return code


def run_check() -> int:
    store = Store(config.DATA_DIR)
    client = ApiClient()
    exit_code = 0
    changed = False
    try:
        now = now_iso()
        try:
            crawl = fetch_all_offers(client, store.state.get("taille_page"))
        except (ApiError, ApiBlocked) as exc:
            return handle_failure(store, str(exc))
        log(f"{len(crawl.offers)} offres récupérées sur {crawl.total} annoncées "
            f"({crawl.pages} page(s), mode {client.mode}).")
        for err in crawl.errors:
            log(f"⚠️  {err}")
        if not crawl.offers:
            return handle_failure(store, "; ".join(crawl.errors) or "liste vide")

        first_run = not store.state.get("initialise")
        counts = {"new": 0, "new_zone": 0, "new_alert": 0, "removed": 0, "relisted": 0, "modified": 0}

        # 1. offres présentes
        for guid, offer in crawl.offers.items():
            rec = store.records.get(guid)
            in_zone = zone_of_offer(offer)
            alert = wants_alert(offer)
            if rec is None:
                store.new_record(offer, now, initial=first_run, in_zone=alert)
                counts["new"] += 1
                counts["new_zone"] += in_zone
                counts["new_alert"] += alert
                if in_zone and not alert and not first_run:
                    log(f"Nouvelle offre dans la zone mais au-dessus de {fmt(ALERT_MAX_RENT, 'money')} : "
                        f"pas d'alerte ({title({'liste': offer})} — {fmt(offer_rent(offer), 'money')}).")
                continue
            if store.update_listing(rec, offer, now):
                counts["modified"] += 1
            if rec["statut"] == "retiree":
                previous = rec.get("retrait_detecte")
                rec.update(statut="en_ligne", absences=0, absente_depuis=None, retrait_detecte=None,
                           remises_en_ligne=(rec.get("remises_en_ligne") or 0) + 1)
                if rec.get("fiche"):
                    rec["fiche_statut"] = "a_rafraichir"
                if alert and config.ALERT_ON_RELISTING:
                    rec["alerte"] = "a_envoyer_remise"
                store.event(now, "Remise en ligne", rec, f"retirée le {fmt_dt(previous)}")
                store.touch(guid)
                counts["relisted"] += 1
            elif rec.get("absences"):
                rec.update(absences=0, absente_depuis=None)  # absence passagère : annulée
                store.touch(guid)

        # 2. offres disparues (seulement si la liste est complète, pour éviter les faux retraits)
        if crawl.complete:
            for rec in store.records.values():
                if rec["statut"] != "en_ligne" or rec["guid"] in crawl.offers:
                    continue
                rec["absences"] = (rec.get("absences") or 0) + 1
                if rec["absences"] == 1:
                    rec["absente_depuis"] = now
                    rec["derniere_presence"] = store.state.get("dernier_passage_complet") or rec["premiere_detection"]
                if rec["absences"] >= config.REMOVAL_CONFIRM_RUNS:
                    rec.update(statut="retiree", retrait_detecte=rec["absente_depuis"])
                    if rec.get("alerte", "") and str(rec["alerte"]).startswith("a_envoyer"):
                        rec["alerte"] = "annulee_retrait"
                    store.event(rec["absente_depuis"], "Retrait", rec,
                                f"dernière présence le {fmt_dt(rec['derniere_presence'])}")
                    counts["removed"] += 1
                store.touch(rec["guid"])
            store.set_state(dernier_passage_complet=now)
        else:
            log(f"⚠️  Liste incomplète ({len(crawl.offers)}/{crawl.total}) : retraits non évalués ce passage-ci.")

        store.set_state(dernier_passage=now, passages_vides=0, alerte_vide_envoyee=False, taille_page=crawl.size)
        if first_run:
            store.set_state(initialise=True, debut_suivi=now)
            store.event(now, "Démarrage du suivi", None, f"{len(crawl.offers)} offres déjà en ligne enregistrées")

        # 3. fiches détaillées
        fresh: dict[str, dict] = {}
        queue = [r for r in store.records.values() if r["statut"] == "en_ligne"
                 and (r.get("fiche") is None or r.get("fiche_statut") == "a_rafraichir")
                 and (r.get("fiche_essais") or 0) < 3]
        queue.sort(key=lambda r: r.get("premiere_detection") or "", reverse=True)  # plus récentes d'abord…
        queue.sort(key=lambda r: (not str(r.get("alerte") or "").startswith("a_envoyer"), not is_in_zone(r)))
        must = sum(1 for r in queue if str(r.get("alerte") or "").startswith("a_envoyer"))
        budget = max(config.MAX_DETAILS_PER_RUN, must)
        ok = 0
        for rec in queue[:budget]:
            time.sleep(config.REQUEST_DELAY)
            data, status = fetch_detail(client, rec["guid"])
            if data:
                fresh[rec["guid"]] = data
                rec.update(fiche=clean(data), fiche_statut="ok", fiche_recuperee_le=now, fiche_essais=0)
                ok += 1
            else:
                rec.update(fiche_statut=status, fiche_essais=(rec.get("fiche_essais") or 0) + 1)
            store.touch(rec["guid"])
        if queue:
            log(f"Fiches détaillées : {ok}/{min(len(queue), budget)} récupérées, "
                f"{max(0, len(queue) - budget)} en attente pour les prochains passages.")

        changed |= store.save()

        # 4. alertes
        pending = sorted((r for r in store.records.values() if r["statut"] == "en_ligne"
                          and str(r.get("alerte") or "").startswith("a_envoyer")),
                         key=lambda r: r["premiere_detection"])
        sent = 0
        if pending and not config.EMAIL_ENABLED:
            log(f"⚠️  {len(pending)} alerte(s) en attente mais l'e-mail n'est pas configuré (SMTP_HOST, EMAIL_TO).")
        elif pending:
            import alertes
            for rec in pending[: config.MAX_ALERTS_PER_RUN]:
                raw = fresh.get(rec["guid"])
                if raw is None:  # les liens des photos expirent : on relit la fiche pour en avoir des frais
                    raw, _ = fetch_detail(client, rec["guid"])
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
        if first_run:
            log(f"Premier passage : {counts['new']} offres enregistrées (dont {counts['new_zone']} dans la zone), "
                "sans alerte. Les prochaines nouveautés de la zone déclencheront un e-mail.")
        else:
            log(f"Bilan : {counts['new']} nouvelle(s) (dont {counts['new_zone']} dans la zone, "
                f"{counts['new_alert']} sous le plafond de loyer), "
                f"{counts['removed']} retrait(s), {counts['relisted']} remise(s) en ligne, "
                f"{counts['modified']} modification(s) · {sent} alerte(s) envoyée(s) · {online} offres en ligne.")
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Veille des offres Action Logement")
    parser.add_argument("--test-email", action="store_true", help="envoie un exemple d'alerte")
    parser.add_argument("--export", action="store_true", help="régénère seulement le fichier Excel")
    args = parser.parse_args()
    if args.test_email:
        return run_test_email()
    if args.export:
        return run_export()
    return run_check()


if __name__ == "__main__":
    sys.exit(main())
