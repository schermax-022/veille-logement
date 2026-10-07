"""Description de toutes les informations d'une offre, partagée par l'export Excel et les e-mails.

Chaque offre enregistrée contient la version « liste » (résultat de recherche) et, quand elle a pu
être récupérée, la « fiche » détaillée. Chaque champ ci-dessous va chercher sa valeur dans la fiche,
puis dans la liste.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Callable
from zoneinfo import ZoneInfo

from config import DETAIL_URL, in_alert_zone

PARIS = ZoneInfo("Europe/Paris")

ENUM_FR = {
    "INDIVIDUAL": "Individuel", "COLLECTIVE": "Collectif", "SHARED": "Partagé", "NONE": "Aucun",
    "UNKNOWN": "Non communiqué", "YES": "Oui", "NO": "Non", "PRESENT": "Oui", "ABSENT": "Non",
    "RENT_STANDARD": "Loyer standard", "IDENTIFIED": "Logement identifié", "GENERIC": "Logement générique",
}


# ---------------------------------------------------------------------------------------------
# Accès aux valeurs
# ---------------------------------------------------------------------------------------------

def dig(obj: Any, path: str) -> Any:
    for part in path.split("."):
        if not isinstance(obj, dict):
            return None
        obj = obj.get(part)
    return obj


def pick(rec: dict, path: str) -> Any:
    """Valeur dans la fiche détaillée, sinon dans la liste."""
    for src in ("fiche", "liste"):
        value = dig(rec.get(src) or {}, path)
        if value not in (None, "", [], {}):
            return value
    return None


def parse_date(value: Any) -> date | None:
    if not value:
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(value))
    return date(int(m[1]), int(m[2]), int(m[3])) if m else None


def parse_dt(value: Any) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value)).astimezone(PARIS)
    except ValueError:
        return None


def labels(value: Any) -> str | None:
    """Liste d'objets {label: …} (services, annexes…) → texte."""
    if not value:
        return None
    if isinstance(value, dict):
        value = [value]
    out = []
    for item in value if isinstance(value, list) else [value]:
        if isinstance(item, dict):
            out.append(str(item.get("commercialLabel") or item.get("label") or item.get("name") or item.get("code") or ""))
        else:
            out.append(ENUM_FR.get(str(item), str(item)))
    out = [o for o in out if o]
    return " ; ".join(out) or None


def enum(value: Any) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, dict):
        return labels(value)
    return ENUM_FR.get(str(value), str(value))


def populations(value: Any) -> str | None:
    """Public visé : [{label: 'Âge', type: 'INTEGER_INTERVAL', intStartValue: 18, intEndValue: 29}] → texte."""
    if not value:
        return None
    out = []
    for p in value:
        if not isinstance(p, dict):
            continue
        label = p.get("label") or p.get("code") or "Critère"
        kind = (p.get("type") or "").upper()
        start, end = p.get("intStartValue"), p.get("intEndValue")
        unit = " ans" if (p.get("group") or p.get("code") or "").upper().startswith("AGE") else ""
        if "INTERVAL" in kind or start is not None or end is not None:
            if start is not None and end is not None:
                out.append(f"{label} : de {start} à {end}{unit}")
            elif start is not None:
                out.append(f"{label} : à partir de {start}{unit}")
            elif end is not None:
                out.append(f"{label} : jusqu'à {end}{unit}")
        elif p.get("boolValue") is not None:
            out.append(f"{label} : {'oui' if p['boolValue'] else 'non'}")
        elif p.get("listValues"):
            out.append(f"{label} : {labels(p['listValues'])}")
        else:
            out.append(str(label))
    return " ; ".join(out) or None


def municipality_code(rec: dict) -> str | None:
    return pick(rec, "lodging.municipality.code")


def department(rec: dict) -> str | None:
    code = municipality_code(rec) or ""
    if code.startswith("97"):
        return code[:3]
    return code[:2] or None


def is_in_zone(rec: dict) -> bool:
    return in_alert_zone(municipality_code(rec))


def photo_urls(rec: dict) -> list[str]:
    pics = (rec.get("fiche") or {}).get("pictures") or (rec.get("liste") or {}).get("pictures") or []
    pics = sorted([p for p in pics if isinstance(p, dict) and p.get("url")], key=lambda p: p.get("ordering") or 0)
    return [p["url"] for p in pics]


def detail_url(rec: dict) -> str:
    return DETAIL_URL.format(guid=rec["guid"])


def positive(value: Any) -> Any:
    return value if isinstance(value, (int, float)) and value > 0 else None


# ---------------------------------------------------------------------------------------------
# Liste des champs
# ---------------------------------------------------------------------------------------------

@dataclass
class Field:
    section: str
    label: str
    get: Callable[[dict], Any]
    kind: str = "text"         # text, money, area, int, date, datetime, bool, longtext, link, formula
    paths: tuple = ()          # chemins bruts couverts (pour repérer les « autres champs »)
    width: int = 16
    in_email: bool = True


def P(path: str) -> Callable[[dict], Any]:
    return lambda rec: pick(rec, path)


def F(section, label, path=None, kind="text", get=None, width=16, in_email=True, extra_paths=()):
    paths = ((path,) if path else ()) + tuple(extra_paths)
    return Field(section, label, get or P(path), kind, paths, width, in_email)


STATUS_FR = {"en_ligne": "En ligne", "retiree": "Retirée"}

FIELDS: list[Field] = [
    # Suivi
    F("Suivi", "Statut", get=lambda r: STATUS_FR.get(r.get("statut"), r.get("statut")), width=11, in_email=False),
    F("Suivi", "Mise en ligne détectée le", kind="datetime", get=lambda r: parse_dt(r.get("premiere_detection")),
      width=18, in_email=False),
    F("Suivi", "Déjà en ligne au démarrage du suivi", kind="bool", get=lambda r: r.get("deja_en_ligne_au_demarrage"),
      width=14, in_email=False),
    F("Suivi", "Publiée le (site)", "startDateOfPublication", "date", width=13),
    F("Suivi", "Fin de publication prévue", "endDateOfPublication", "date", width=13),
    F("Suivi", "Vue en ligne pour la dernière fois", kind="datetime", get=lambda r: parse_dt(r.get("derniere_presence")),
      width=18, in_email=False),
    F("Suivi", "Retrait détecté le", kind="datetime", get=lambda r: parse_dt(r.get("retrait_detecte")),
      width=18, in_email=False),
    F("Suivi", "Durée en ligne (jours)", kind="formula", get=lambda r: None, width=12, in_email=False),
    F("Suivi", "Remises en ligne", kind="int", get=lambda r: r.get("remises_en_ligne") or 0, width=10, in_email=False),
    F("Suivi", "Zone d'alerte", kind="bool", get=is_in_zone, width=10, in_email=False),
    # Localisation
    F("Localisation", "Ville", "lodging.municipality.name", width=26),
    F("Localisation", "Code postal", "lodging.municipality.postcode", width=10),
    F("Localisation", "Code INSEE", "lodging.municipality.code", width=10, in_email=False),
    F("Localisation", "Département", get=department, width=8, in_email=False),
    # Loyer et frais
    F("Loyer et frais", "Loyer affiché (€/mois)", "totalRentAmountBaseBound", "money", width=12),
    F("Loyer et frais", "Loyer maximum (€/mois)", "totalRentAmountMax", "money", width=12),
    F("Loyer et frais", "Redevance (€/mois)", "housingFeeBaseBound", "money", width=12),
    F("Loyer et frais", "Charges (€/mois)", "rentChargeAmount", "money", width=11),
    F("Loyer et frais", "Charges annexes incluses", "rentAnnexChargesIncluded", "bool", width=11),
    F("Loyer et frais", "Charges externes estimées (€)", "estimatedExternalCharges", "money", width=12),
    F("Loyer et frais", "Stationnement (€/mois)", "rentParkingAmount", "money", width=12),
    F("Loyer et frais", "Prestations obligatoires (€)", "compulsoryBenefitAmount", "money", width=12),
    F("Loyer et frais", "Loyer progressif (€)", "totalProgressiveRentAmount", "money", width=12),
    F("Loyer et frais", "Dépôt de garantie (€)", "securityDepositAmount", "money", width=12),
    F("Loyer et frais", "Frais d'agence demandés", "areAgencyFeesRequested", "bool", width=10),
    F("Loyer et frais", "Frais d'agence (€)", "agencyFeesAmount", "money", width=11),
    F("Loyer et frais", "Loyer de référence majoré", "isOfferRentReferenceIncreased", "bool", width=11),
    F("Loyer et frais", "Loyer de référence majoré (€)", "rentReferenceIncreased", "money", width=12),
    F("Loyer et frais", "Type de loyer", get=lambda r: enum(pick(r, "rentType")), width=14, extra_paths=("rentType",)),
    F("Loyer et frais", "Paiement", get=lambda r: labels(pick(r, "paymentType")), width=10, extra_paths=("paymentType",)),
    # Logement
    F("Logement", "Catégorie", get=lambda r: labels(pick(r, "lodging.lodgingCategory")), width=13,
      extra_paths=("lodging.lodgingCategory",)),
    F("Logement", "Typologie", get=lambda r: labels(pick(r, "lodging.lodgingTypology")), width=9,
      extra_paths=("lodging.lodgingTypology",)),
    F("Logement", "Surface (m²)", "lodging.area", "area", width=9),
    F("Logement", "Pièces", get=lambda r: positive(pick(r, "lodging.numberOfRoomsHousing")), kind="int", width=7,
      extra_paths=("lodging.numberOfRoomsHousing",)),
    F("Logement", "Étage", get=lambda r: labels(pick(r, "lodging.floor")), width=7, extra_paths=("lodging.floor",)),
    F("Logement", "Meublé", "lodging.isFurnished", "bool", width=8),
    F("Logement", "Ascenseur", get=lambda r: enum(pick(r, "lodging.elevatorPresence")), width=10,
      extra_paths=("lodging.elevatorPresence",)),
    F("Logement", "Type d'appartement", get=lambda r: enum(pick(r, "lodging.apartmentType")), width=12,
      extra_paths=("lodging.apartmentType",)),
    F("Logement", "Cuisine", get=lambda r: enum(pick(r, "lodging.kitchenType")), width=11,
      extra_paths=("lodging.kitchenType",)),
    F("Logement", "Cuisine équipée", "lodging.isKitchenEquipped", "bool", width=9),
    F("Logement", "Salle de bains", get=lambda r: enum(pick(r, "lodging.bathroomType")), width=11,
      extra_paths=("lodging.bathroomType",)),
    F("Logement", "WC", get=lambda r: enum(pick(r, "lodging.toiletType")), width=11, extra_paths=("lodging.toiletType",)),
    F("Logement", "Annexes", get=lambda r: labels(pick(r, "lodging.lodgingAnnexTypes")), width=16,
      extra_paths=("lodging.lodgingAnnexTypes",)),
    F("Logement", "Extérieurs", get=lambda r: labels(pick(r, "lodging.outdoorTypes")), width=16,
      extra_paths=("lodging.outdoorTypes",)),
    F("Logement", "Stationnement", get=lambda r: labels(pick(r, "lodging.parkingTypes")), width=14,
      extra_paths=("lodging.parkingTypes",)),
    F("Logement", "Places de stationnement", "lodging.parkingCapacity", "int", width=10),
    F("Logement", "Stationnement obligatoire", "lodging.isParkingMandatory", "bool", width=11),
    F("Logement", "Accessibilité PMR", get=lambda r: enum(pick(r, "lodging.prmMode")), width=11,
      extra_paths=("lodging.prmMode",)),
    F("Logement", "Logement neuf", "isNewLodging", "bool", width=9),
    F("Logement", "Logement réhabilité", "isLodgingRehabilitated", "bool", width=10),
    F("Logement", "Première mise en location", "isLodgingFirstTimeAvailable", "bool", width=11),
    F("Logement", "Disponible le", "availableAt", "date", width=12),
    # Énergie et équipements collectifs
    F("Énergie", "Chauffage", get=lambda r: labels(pick(r, "lodging.heatingMode")), width=20,
      extra_paths=("lodging.heatingMode",)),
    F("Énergie", "Eau chaude", get=lambda r: enum(pick(r, "lodging.hotWaterType")), width=11,
      extra_paths=("lodging.hotWaterType",)),
    F("Énergie", "Eau froide", get=lambda r: enum(pick(r, "lodging.coldWaterType")), width=11,
      extra_paths=("lodging.coldWaterType",)),
    F("Énergie", "Électricité", get=lambda r: enum(pick(r, "lodging.electricityType")), width=11,
      extra_paths=("lodging.electricityType",)),
    F("Énergie", "DPE (énergie)", get=lambda r: labels(pick(r, "lodging.energyPerformanceCertificate")), width=11,
      extra_paths=("lodging.energyPerformanceCertificate",)),
    F("Énergie", "GES (climat)", get=lambda r: labels(pick(r, "lodging.greenhouseGasCertificate")), width=11,
      extra_paths=("lodging.greenhouseGasCertificate",)),
    F("Énergie", "Dépense énergétique annuelle (€)", get=lambda r: positive(pick(r, "lodging.annualEnergyExpenditure")),
      kind="money", width=12, extra_paths=("lodging.annualEnergyExpenditure",)),
    # Résidence
    F("Résidence", "Résidence", "residency.name", width=28),
    F("Résidence", "Type de résidence", get=lambda r: labels(pick(r, "residency.type")), width=22,
      extra_paths=("residency.type",)),
    F("Résidence", "Patrimoine", get=lambda r: labels(pick(r, "residency.type.product")), width=18,
      extra_paths=("residency.type.product",)),
    F("Résidence", "Règles de vie", "residency.livingRules", width=24),
    F("Résidence", "Services", get=lambda r: labels(pick(r, "residency.services")), width=20,
      extra_paths=("residency.services",)),
    F("Résidence", "Modes de vie", get=lambda r: labels(pick(r, "residency.lifeStyles")), width=16,
      extra_paths=("residency.lifeStyles",)),
    F("Résidence", "Garanties exigées", get=lambda r: labels(pick(r, "residency.requiredWarrantyTypes")), width=16,
      extra_paths=("residency.requiredWarrantyTypes",)),
    F("Résidence", "Année de construction", "residency.constructionYear", "int", width=10),
    F("Résidence", "Livrée le", "residency.deliveredAt", "date", width=12),
    F("Résidence", "Dernière réhabilitation", "residency.lastRehabilitatedDate", "date", width=12),
    # Conditions d'accès
    F("Conditions", "Public visé", get=lambda r: populations(pick(r, "offerTargetPopulations")), width=22,
      extra_paths=("offerTargetPopulations",)),
    F("Conditions", "Numéro unique de demande requis", "isDemandUniqueNumberRequired", "bool", width=12),
    F("Conditions", "Catégorie d'offre", get=lambda r: enum(pick(r, "offerCategory")), width=12,
      extra_paths=("offerCategory",)),
    F("Conditions", "Code d'orientation", get=lambda r: enum(pick(r, "orientationCode")), width=12,
      extra_paths=("orientationCode",)),
    F("Conditions", "Offre disponible (selon la fiche)", "isOfferAvailable", "bool", width=12),
    # Textes
    F("Description", "Description de l'offre", "description", "longtext", width=60),
    F("Description", "Description du logement", "lodging.description", "longtext", width=50),
    F("Description", "Photos", get=lambda r: len(photo_urls(r)), kind="int", width=7, in_email=False,
      extra_paths=("pictures",)),
    # Références
    F("Références", "Lien vers l'annonce", get=detail_url, kind="link", width=14, in_email=False),
    F("Références", "Référence de l'offre", "reference", width=11),
    F("Références", "Numéro d'offre", "offerNumber", "int", width=8, in_email=False),
    F("Références", "Référence du logement", "lodging.reference", width=11),
    F("Références", "Référence de la résidence", "residency.reference", width=11, in_email=False),
    F("Références", "Identifiant (GUID)", "guid", width=38, in_email=False),
]

COVERED_PATHS = {p for f in FIELDS for p in f.paths}
# Identifiants internes et doublons sans intérêt pour l'utilisateur
_SKIP_LEAF = {"id", "ordering", "guid", "dimension", "urlContentType", "url", "code", "commercialLabel"}


def extra_fields(rec: dict) -> dict[str, Any]:
    """Tous les autres champs bruts renvoyés par le site, pour ne rien perdre."""
    out: dict[str, Any] = {}

    def walk(obj: Any, path: str) -> None:
        if any(path == c or path.startswith(c + ".") for c in COVERED_PATHS):
            return
        if isinstance(obj, dict):
            if "label" in obj and path:          # objet « énuméré » : on garde son libellé
                out.setdefault(path, labels(obj))
                return
            for k, v in obj.items():
                if k in _SKIP_LEAF:
                    continue
                walk(v, f"{path}.{k}" if path else k)
        elif isinstance(obj, list):
            if obj and all(isinstance(x, dict) for x in obj):
                out.setdefault(path, labels(obj))
            elif obj:
                out.setdefault(path, " ; ".join(map(str, obj)))
        elif obj is not None and obj != "":
            out.setdefault(path, obj)

    for src in ("fiche", "liste"):
        walk(rec.get(src) or {}, "")
    for path in ("demandGuid", "candidacy", "demandUniqueNumberCorrespondenceCode", "isRecent"):
        out.pop(path, None)
    return out


# ---------------------------------------------------------------------------------------------
# Mise en forme française (e-mails, historique)
# ---------------------------------------------------------------------------------------------

def fr_number(value: float, decimals: int = 2) -> str:
    s = f"{value:,.{decimals}f}".replace(",", " ").replace(".", ",")
    return s[:-3] if decimals == 2 and s.endswith(",00") else s


def fmt(value: Any, kind: str) -> str | None:
    if value in (None, "", []):
        return None
    if kind == "money":
        return f"{fr_number(float(value))} €"
    if kind == "area":
        return f"{fr_number(float(value), 2).rstrip('0').rstrip(',')} m²"
    if kind == "date":
        d = parse_date(value) if not isinstance(value, date) else value
        return d.strftime("%d/%m/%Y") if d else str(value)
    if kind == "datetime":
        return value.strftime("%d/%m/%Y à %H:%M") if isinstance(value, datetime) else str(value)
    if kind == "bool":
        return "Oui" if value else "Non"
    return str(value)


def title(rec: dict) -> str:
    """Titre au format du site : « Location - T2 - 45.6 m² - Paris 18e Arrondissement (75018) »."""
    parts = ["Location"]
    typ = labels(pick(rec, "lodging.lodgingTypology"))
    if typ:
        parts.append(typ)
    area = pick(rec, "lodging.area")
    if area:
        parts.append(fmt(area, "area"))
    city = pick(rec, "lodging.municipality.name")
    cp = pick(rec, "lodging.municipality.postcode")
    if city:
        parts.append(f"{city} ({cp})" if cp else city)
    return " - ".join(parts)


def short_city(rec: dict) -> str:
    city = pick(rec, "lodging.municipality.name") or ""
    m = re.match(r"Paris (\d+)(?:er|e)", city)
    return f"Paris {m[1]}e" if m and m[1] != "1" else ("Paris 1er" if m else city)


def rent_value(rec: dict) -> Any:
    for path in ("totalRentAmountBaseBound", "housingFeeBaseBound", "totalRentAmountMax"):
        v = pick(rec, path)
        if v:
            return v
    return None
