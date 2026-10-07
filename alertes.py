"""E-mails d'alerte : une reproduction de l'annonce, photos comprises."""

from __future__ import annotations

import re
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from html import escape

import config
from annonce import (FIELDS, PARIS, detail_url, fmt, labels, parse_dt, pick, rent_value, short_city,
                     title)

ACCENT = "#e2001a"
INK = "#1b2430"
MUTED = "#5f6875"
LINE = "#e3e6ea"

# Champs déjà mis en avant en haut de l'e-mail (pas répétés dans les tableaux)
HIGHLIGHTED = {"Ville", "Code postal", "Typologie", "Surface (m²)", "Loyer affiché (€/mois)", "Disponible le",
               "Description de l'offre", "Description du logement", "Résidence", "Étage", "Meublé", "Catégorie"}
EMAIL_SECTIONS = ["Loyer et frais", "Logement", "Énergie", "Résidence", "Conditions", "Suivi", "Références"]


def h(value) -> str:
    return escape(str(value), quote=True)


def _image_type(data: bytes, hint: str) -> tuple[str, str]:
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "image", "png"
    if data[:3] == b"\xff\xd8\xff":
        return "image", "jpeg"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image", "webp"
    if data[:4] == b"GIF8":
        return "image", "gif"
    if hint.startswith("image/"):
        return "image", hint.split("/", 1)[1].split(";")[0]
    return "image", "jpeg"


def download_photos(client, urls: list[str]) -> list[tuple[bytes, str, str]]:
    photos = []
    for url in urls[: config.MAX_PHOTOS]:
        try:
            data, ctype = client.get_bytes(url)
        except Exception as exc:
            print(f"   photo non téléchargée ({exc})", flush=True)
            continue
        if len(data) > 200:
            maintype, subtype = _image_type(data, ctype or "")
            photos.append((data, maintype, subtype))
    return photos


def _rows(rec: dict) -> dict[str, list[tuple[str, str]]]:
    sections: dict[str, list[tuple[str, str]]] = {}
    for f in FIELDS:
        if not f.in_email or f.label in HIGHLIGHTED or f.section not in EMAIL_SECTIONS:
            continue
        text = fmt(f.get(rec), f.kind)
        if text:
            sections.setdefault(f.section, []).append((f.label.replace(" (€/mois)", "").replace(" (€)", ""), text))
    return sections


def _mostly_contained(small: str, big: str) -> bool:
    words = lambda t: set(re.findall(r"\w{3,}", t.lower()))
    a, b = words(small), words(big)
    return bool(a) and len(a & b) / len(a) >= 0.8


def _facts(rec: dict) -> list[tuple[str, str]]:
    facts = [
        ("Type", labels(pick(rec, "lodging.lodgingTypology"))),
        ("Surface", fmt(pick(rec, "lodging.area"), "area")),
        ("Étage", labels(pick(rec, "lodging.floor"))),
        ("Meublé", fmt(pick(rec, "lodging.isFurnished"), "bool")),
        ("Disponible le", fmt(pick(rec, "availableAt"), "date")),
        ("Catégorie", labels(pick(rec, "lodging.lodgingCategory"))),
    ]
    return [(k, v) for k, v in facts if v]


def subject(rec: dict, relisted: bool = False) -> str:
    bits = [short_city(rec), labels(pick(rec, "lodging.lodgingTypology")), fmt(pick(rec, "lodging.area"), "area")]
    rent = rent_value(rec)
    if rent:
        bits.append(f"{fmt(rent, 'money')}/mois")
    prefix = "🔁 De nouveau en ligne" if relisted else "🏠 Nouvelle annonce"
    return f"{prefix} · " + " · ".join(b for b in bits if b)


def render(rec: dict, photo_cids: list[str], relisted: bool = False) -> tuple[str, str]:
    """Renvoie (texte, html) de l'alerte."""
    url = detail_url(rec)
    detected = parse_dt(rec.get("premiere_detection")) or datetime.now(PARIS)
    kicker = "De nouveau en ligne" if relisted else "Nouvelle annonce"
    when = f"détectée le {detected.strftime('%d/%m/%Y à %H:%M')}"
    rent = rent_value(rec)
    residence = pick(rec, "residency.name")
    tags = [labels(pick(rec, "residency.type")), labels(pick(rec, "residency.type.product"))]
    tags = [t for t in tags if t]
    facts = _facts(rec)
    sections = _rows(rec)
    desc_offer = pick(rec, "description")
    desc_lodging = pick(rec, "lodging.description")
    if desc_offer and desc_lodging and _mostly_contained(desc_lodging, desc_offer):
        desc_lodging = None
    no_detail = not rec.get("fiche")

    # ---- texte brut ----
    lines = [f"{kicker} — {when}", "", title(rec)]
    if residence:
        lines.append(residence + (f" ({', '.join(tags)})" if tags else ""))
    if rent:
        lines.append(f"Loyer affiché : {fmt(rent, 'money')} / mois")
    lines += [f"{k} : {v}" for k, v in facts]
    def text_section(section):
        return ["", section.upper()] + [f"  {k} : {v}" for k, v in sections[section]] if sections.get(section) else []
    for section in ("Loyer et frais", "Logement"):
        lines += text_section(section)
    for label, text in (("DESCRIPTION", desc_offer), ("LE LOGEMENT", desc_lodging)):
        if text:
            lines += ["", label, text.strip()]
    for section in EMAIL_SECTIONS:
        if section not in ("Loyer et frais", "Logement"):
            lines += text_section(section)
    lines += ["", f"Voir l'annonce et candidater : {url}"]
    text_body = "\n".join(lines) + "\n"

    # ---- HTML ----
    def table(rows: list[tuple[str, str]]) -> str:
        return "".join(
            f'<tr><td style="padding:7px 0;border-top:1px solid {LINE};color:{MUTED};font-size:13px;width:44%;'
            f'vertical-align:top">{h(k)}</td><td style="padding:7px 0;border-top:1px solid {LINE};color:{INK};'
            f'font-size:14px;vertical-align:top">{h(v)}</td></tr>' for k, v in rows)

    photos_html = ""
    if photo_cids:
        photos_html = (f'<tr><td style="padding:0 0 6px 0"><img src="cid:{photo_cids[0]}" width="560" alt="Photo 1" '
                       f'style="display:block;width:100%;max-width:560px;height:auto;border-radius:8px"></td></tr>')
        thumbs = photo_cids[1:4]
        if thumbs:
            cells = "".join(
                f'<td width="{100 // len(thumbs)}%" style="padding:0 {3 if j < len(thumbs) - 1 else 0}px 0 0">'
                f'<img src="cid:{cid}" alt="Photo {j + 2}" style="display:block;width:100%;height:auto;'
                f'border-radius:6px"></td>' for j, cid in enumerate(thumbs))
            photos_html += (f'<tr><td style="padding:0 0 8px 0"><table role="presentation" width="100%" '
                            f'cellpadding="0" cellspacing="0"><tr>{cells}</tr></table></td></tr>')
    facts_html = "".join(
        f'<td style="padding:10px 12px;background:#f6f7f9;border-radius:8px;text-align:center">'
        f'<div style="font-size:11px;color:{MUTED};text-transform:uppercase;letter-spacing:.04em">{h(k)}</div>'
        f'<div style="font-size:15px;font-weight:600;color:{INK};margin-top:2px">{h(v)}</div></td>'
        f'<td style="width:6px"></td>' for k, v in facts[:4])
    facts_rest = facts[4:]
    top_html, bottom_html = "", ""
    for section in EMAIL_SECTIONS:
        rows = sections.get(section)
        if section == "Logement":
            rows = facts_rest + (rows or [])
        if not rows:
            continue
        block = (
            f'<tr><td style="padding:22px 0 4px 0;font-size:15px;font-weight:700;color:{INK}">{h(section)}</td></tr>'
            f'<tr><td><table role="presentation" width="100%" cellpadding="0" cellspacing="0">{table(rows)}'
            f'</table></td></tr>')
        if section in ("Loyer et frais", "Logement"):
            top_html += block
        else:
            bottom_html += block
    desc_html = ""
    for label, text in (("Description", desc_offer), ("Le logement", desc_lodging)):
        if text:
            desc_html += (
                f'<tr><td style="padding:22px 0 6px 0;font-size:15px;font-weight:700;color:{INK}">{h(label)}</td></tr>'
                f'<tr><td style="font-size:14px;line-height:1.55;color:{INK};white-space:pre-line">'
                f'{h(text.strip())}</td></tr>')
    warning = ('<tr><td style="padding:10px 12px;background:#fff6e5;border-radius:8px;font-size:13px;color:#7a4b00">'
               'La fiche détaillée n\'a pas pu être lue : seules les informations de la liste sont affichées.'
               '</td></tr>' if no_detail else "")
    preheader = " · ".join(v for _, v in facts[:3]) + (f" · {fmt(rent, 'money')}" if rent else "")

    html_body = f"""<!doctype html><html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#eef0f3">
<div style="display:none;max-height:0;overflow:hidden">{h(preheader)}</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef0f3">
<tr><td align="center" style="padding:20px 10px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
 style="max-width:600px;background:#ffffff;border-radius:12px;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif">
<tr><td style="padding:22px 20px 26px 20px">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0">
  <tr><td style="font-size:12px;color:{ACCENT};font-weight:700;text-transform:uppercase;letter-spacing:.05em">
    {h(kicker)} <span style="color:{MUTED};font-weight:400;text-transform:none;letter-spacing:0">· {h(when)}</span></td></tr>
  <tr><td style="padding:6px 0 4px 0;font-size:21px;line-height:1.3;font-weight:700;color:{INK}">{h(title(rec))}</td></tr>
  {f'<tr><td style="padding:0 0 4px 0;font-size:14px;color:{INK}">{h(residence)}</td></tr>' if residence else ''}
  {f'<tr><td style="padding:0 0 14px 0;font-size:13px;color:{MUTED}">{h(" · ".join(tags))}</td></tr>' if tags else ''}
  {photos_html}
  {f'<tr><td style="padding:10px 0 2px 0"><span style="font-size:26px;font-weight:800;color:{INK}">{h(fmt(rent, "money"))}</span> <span style="font-size:14px;color:{MUTED}">/ mois · loyer affiché</span></td></tr>' if rent else ''}
  <tr><td style="padding:12px 0 0 0"><table role="presentation" cellpadding="0" cellspacing="0"><tr>{facts_html}</tr></table></td></tr>
  <tr><td style="padding:20px 0 0 0"><a href="{h(url)}" style="display:inline-block;padding:12px 20px;background:{ACCENT};
    color:#ffffff;text-decoration:none;border-radius:8px;font-size:15px;font-weight:700">Voir l'annonce et candidater</a></td></tr>
  {warning}
  {top_html}
  {desc_html}
  {bottom_html}
  <tr><td style="padding:26px 0 0 0;font-size:12px;color:{MUTED}">
    Alerte envoyée par votre veille Action Logement · <a href="{h(url)}" style="color:{MUTED}">{h(url)}</a></td></tr>
</table></td></tr></table></td></tr></table></body></html>"""
    return text_body, html_body


# ---------------------------------------------------------------------------------
# Envoi
# ---------------------------------------------------------------------------------

def _send(msg: EmailMessage) -> None:
    ctx = ssl.create_default_context()
    if config.SMTP_SECURITY == "ssl":
        server = smtplib.SMTP_SSL(config.SMTP_HOST, config.SMTP_PORT, context=ctx, timeout=60)
    else:
        server = smtplib.SMTP(config.SMTP_HOST, config.SMTP_PORT, timeout=60)
    with server as s:
        if config.SMTP_SECURITY == "starttls":
            s.starttls(context=ctx)
        if config.SMTP_USER:
            s.login(config.SMTP_USER, config.SMTP_PASSWORD)
        s.send_message(msg, to_addrs=config.EMAIL_TO)


def _base_message(subj: str) -> EmailMessage:
    sender = config.EMAIL_FROM
    msg = EmailMessage()
    msg["Subject"] = subj
    msg["From"] = formataddr(("Veille Action Logement", sender)) if "<" not in sender else sender
    msg["To"] = ", ".join(config.EMAIL_TO)
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(sender.rsplit("@", 1)[-1].strip("> ") or "veille.local"))
    return msg


def send_rendered(render_fn, photos: list[tuple[bytes, str, str]]) -> None:
    """Envoie un e-mail mis en forme ; render_fn(cids) renvoie (objet, texte, html)."""
    cids = [make_msgid(domain="veille.local")[1:-1] for _ in photos]
    subj, text_body, html_body = render_fn(cids)
    msg = _base_message(subj)
    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")
    html_part = msg.get_payload()[1]
    for cid, (data, maintype, subtype) in zip(cids, photos):
        html_part.add_related(data, maintype=maintype, subtype=subtype, cid=f"<{cid}>",
                              filename=f"photo.{subtype}")
    _send(msg)


def send_offer_alert(rec: dict, photos: list[tuple[bytes, str, str]], relisted: bool = False) -> None:
    send_rendered(lambda cids: (subject(rec, relisted), *render(rec, cids, relisted)), photos)


def send_simple(subj: str, body: str, link: str | None = None) -> None:
    msg = _base_message(subj)
    msg.set_content(body + (f"\n\n{link}" if link else ""))
    msg.add_alternative(
        f'<!doctype html><html lang="fr"><body style="font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,'
        f'sans-serif;color:{INK};padding:20px"><p style="font-size:16px;font-weight:700">{h(subj)}</p>'
        f'<p style="font-size:14px;white-space:pre-line">{h(body)}</p>'
        + (f'<p><a href="{h(link)}">{h(link)}</a></p>' if link else "") + "</body></html>", subtype="html")
    _send(msg)
