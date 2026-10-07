"""Logique de suivi commune à tous les sites : mises en ligne, retraits, remises en ligne, échecs."""

from __future__ import annotations

from typing import Callable

import config
from al_api import log
from annonce import fmt, parse_dt
from stockage import Store


def fmt_dt(iso: str | None) -> str:
    dt = parse_dt(iso)
    return dt.strftime("%d/%m/%Y %H:%M") if dt else "—"


def handle_failure(store: Store, reason: str, site: str, site_url: str) -> int:
    """Échec de récupération : on prévient une seule fois après plusieurs échecs d'affilée."""
    n = (store.state.get("passages_vides") or 0) + 1
    store.set_state(passages_vides=n)
    log(f"⚠️  {site} : aucune offre récupérée ({reason}) — {n} passage(s) en échec d'affilée.")
    code = 0
    if n >= config.EMPTY_ALERT_AFTER and not store.state.get("alerte_vide_envoyee"):
        if config.EMAIL_ENABLED:
            import alertes
            try:
                alertes.send_simple(
                    f"⚠️ Veille {site} : plus aucune offre récupérée",
                    f"Depuis {n} vérifications, le script ne récupère plus les offres de {site}.\n\n"
                    f"Dernière erreur : {reason}\n\n"
                    "Si l'erreur parle d'accès refusé ou de page vide, ajoutez la variable USE_BROWSER = true sur GitHub.",
                    site_url, sender_name=f"Veille {site}")
            except Exception as exc:
                log(f"⚠️  E-mail d'avertissement non envoyé : {exc}")
        store.set_state(alerte_vide_envoyee=True)
        code = 1  # un seul échec visible côté GitHub, pas un à chaque passage
    store.save()
    return code


def reconcile(store: Store, offers: dict, now: str, *, complete: bool, total: int,
              in_zone: Callable[[dict], bool], wants_alert: Callable[[dict], bool],
              describe: Callable[[dict], str], rent_of: Callable[[dict], float | None],
              max_rent: float | None) -> dict:
    """Compare la liste du moment avec la base : nouvelles offres, modifications, remises en ligne, retraits."""
    first_run = not store.state.get("initialise")
    counts = {"new": 0, "new_zone": 0, "new_alert": 0, "removed": 0, "relisted": 0, "modified": 0,
              "first_run": first_run}

    # 1. offres présentes
    for guid, offer in offers.items():
        rec = store.records.get(guid)
        zone = in_zone(offer)
        alert = wants_alert(offer)
        if rec is None:
            store.new_record(offer, now, initial=first_run, in_zone=alert)
            counts["new"] += 1
            counts["new_zone"] += zone
            counts["new_alert"] += alert
            if zone and not alert and not first_run:
                log(f"Nouvelle offre dans la zone mais au-dessus de {fmt(max_rent, 'money')} : "
                    f"pas d'alerte ({describe(offer)} — {fmt(rent_of(offer), 'money')}).")
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
    if complete:
        for rec in store.records.values():
            if rec["statut"] != "en_ligne" or rec["guid"] in offers:
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
        log(f"⚠️  Liste incomplète ({len(offers)}/{total}) : retraits non évalués ce passage-ci.")

    store.set_state(dernier_passage=now, passages_vides=0, alerte_vide_envoyee=False)
    if first_run:
        store.set_state(initialise=True, debut_suivi=now)
        store.event(now, "Démarrage du suivi", None, f"{len(offers)} offres déjà en ligne enregistrées")
    return counts


def detail_queue(store: Store, in_zone: Callable[[dict], bool], per_run: int) -> tuple[list, int]:
    """Offres dont il faut lire la fiche : alertes en attente d'abord, puis la zone, puis les plus récentes."""
    queue = [r for r in store.records.values() if r["statut"] == "en_ligne"
             and (r.get("fiche") is None or r.get("fiche_statut") == "a_rafraichir")
             and (r.get("fiche_essais") or 0) < 3]
    queue.sort(key=lambda r: r.get("premiere_detection") or "", reverse=True)
    queue.sort(key=lambda r: (not str(r.get("alerte") or "").startswith("a_envoyer"), not in_zone(r)))
    must = sum(1 for r in queue if str(r.get("alerte") or "").startswith("a_envoyer"))
    return queue, max(per_run, must)


def pending_alerts(store: Store) -> list[dict]:
    return sorted((r for r in store.records.values() if r["statut"] == "en_ligne"
                   and str(r.get("alerte") or "").startswith("a_envoyer")),
                  key=lambda r: r["premiere_detection"])


def bilan(site: str, counts: dict, sent: int, online: int, n_offers: int) -> None:
    if counts["first_run"]:
        log(f"{site} — premier passage : {counts['new']} offres enregistrées (dont {counts['new_zone']} dans la zone), "
            "sans alerte. Les prochaines nouveautés de la zone déclencheront un e-mail.")
    else:
        log(f"{site} — bilan : {counts['new']} nouvelle(s) (dont {counts['new_zone']} dans la zone, "
            f"{counts['new_alert']} sous le plafond de loyer), {counts['removed']} retrait(s), "
            f"{counts['relisted']} remise(s) en ligne, {counts['modified']} modification(s) · "
            f"{sent} alerte(s) envoyée(s) · {online} offres en ligne.")
