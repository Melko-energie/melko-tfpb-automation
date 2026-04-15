import os
from pathlib import Path
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, black
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from backend.config.constants import MELKO_INFO, LOGO_PATH, ARTICLE_CGI
from backend.utils.logger import get_logger

log = get_logger()

# ── Register Calibri font (fallback to Helvetica) ──
FONT = "Helvetica"
FONT_B = "Helvetica-Bold"
FONT_I = "Helvetica-Oblique"
FONT_BI = "Helvetica-BoldOblique"

_fonts_dir = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
_calibri_map = {
    "Calibri": "calibri.ttf",
    "Calibri-Bold": "calibrib.ttf",
    "Calibri-Italic": "calibrii.ttf",
    "Calibri-BoldItalic": "calibriz.ttf",
}
_cal_ok = True
for _name, _file in _calibri_map.items():
    _p = _fonts_dir / _file
    if _p.exists():
        try:
            pdfmetrics.registerFont(TTFont(_name, str(_p)))
        except Exception:
            _cal_ok = False
            break
    else:
        _cal_ok = False
        break

if _cal_ok:
    FONT = "Calibri"
    FONT_B = "Calibri-Bold"
    FONT_I = "Calibri-Italic"
    FONT_BI = "Calibri-BoldItalic"

# ── Colors ──
BLUE_DARK = HexColor("#1F4E79")
BLUE_ACCENT = HexColor("#2AAAE1")
BLUE_BAND_START = HexColor("#3CC0F0")
BLUE_BAND_END = HexColor("#2A8FC7")
GREY_FOOTER = HexColor("#888888")


def _get_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        "MelkoTitle", fontName=FONT_B, fontSize=10,
        textColor=BLUE_DARK, leading=14,
    ))
    styles.add(ParagraphStyle(
        "MelkoBody", fontName=FONT, fontSize=10,
        leading=14, alignment=TA_JUSTIFY, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "MelkoBodyBold", fontName=FONT_B, fontSize=10,
        leading=14, alignment=TA_JUSTIFY, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "MelkoH1", fontName=FONT_B, fontSize=10,
        textColor=BLUE_DARK, leading=14, spaceAfter=8, spaceBefore=14,
        leftIndent=20,
    ))
    styles.add(ParagraphStyle(
        "MelkoH2", fontName=FONT_B, fontSize=10,
        leading=13, spaceAfter=4, spaceBefore=10,
        leftIndent=10,
    ))
    styles.add(ParagraphStyle(
        "MelkoSmall", fontName=FONT, fontSize=7,
        textColor=GREY_FOOTER, leading=9, alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "MelkoObjet", fontName=FONT_B, fontSize=10,
        textColor=BLUE_ACCENT, leading=13, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "MelkoBullet", fontName=FONT, fontSize=10,
        leading=13, leftIndent=30, spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        "MelkoBulletItalic", fontName=FONT_I, fontSize=10,
        leading=13, leftIndent=50, spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        "MelkoSubBullet", fontName=FONT_I, fontSize=10,
        leading=13, leftIndent=60, spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        "MelkoQuoteItem", fontName=FONT_I, fontSize=10,
        leading=13, leftIndent=60, spaceAfter=3,
    ))
    return styles


def _format_money(amount: float) -> str:
    if amount == 0:
        return "0 \u20ac"
    formatted = f"{amount:,.2f}".replace(",", "\u00a0").replace(".", ",")
    return f"{formatted} \u20ac"


def _format_date_fr() -> str:
    days = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
    months = ["janvier", "f\u00e9vrier", "mars", "avril", "mai", "juin",
              "juillet", "ao\u00fbt", "septembre", "octobre", "novembre", "d\u00e9cembre"]
    now = datetime.now()
    return f"le {days[now.weekday()]} {now.day} {months[now.month - 1]} {now.year}"


def _reverse_name(full_name: str) -> str:
    """'Amaury Mongongu' -> 'Mongongu Amaury'"""
    parts = full_name.strip().split()
    if len(parts) == 2:
        return f"{parts[1]} {parts[0]}"
    return full_name


def generate_courrier(
    synthesis: dict,
    commune: str,
    work_type: str,
    output_path: Path,
    demande_num: int = 0,
) -> Path:
    """Generate the courrier PDF for a given segment."""
    log.info(f"Generation courrier PDF : {commune} / {work_type}")

    file_path = output_path / "courrier.pdf"
    file_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(file_path),
        pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=12 * mm, bottomMargin=18 * mm,
    )

    styles = _get_styles()
    story = []

    # ── Logo ──
    if LOGO_PATH.exists():
        logo = Image(str(LOGO_PATH), width=28 * mm, height=28 * mm)
        logo.hAlign = "LEFT"
        story.append(logo)
    story.append(Spacer(1, 4 * mm))

    # ── Header block (sender / recipient) ──
    _add_header_block(story, styles, synthesis, commune)

    # ── Date ──
    story.append(Spacer(1, 8 * mm))
    story.append(Paragraph(f"Fait \u00e0 Paris, {_format_date_fr()},", styles["MelkoBody"]))
    story.append(Spacer(1, 5 * mm))

    # ── Affaire suivie par ──
    reversed_name = _reverse_name(MELKO_INFO["signataire_nom"])
    story.append(Paragraph(
        f'<u>Affaire suivie par</u>\u00a0: {reversed_name}, '
        f'<a href="mailto:{MELKO_INFO["signataire_email"]}" color="#2AAAE1">'
        f'{MELKO_INFO["signataire_email"]}</a> (demande N\u00b0{demande_num})',
        styles["MelkoObjet"]
    ))
    story.append(Spacer(1, 2 * mm))

    # ── Objet ──
    type_labels = {
        "TEE": "Travaux d\u2019\u00e9conomie d\u2019\u00e9nergie dans le cadre d\u2019une r\u00e9novation globale",
        "PMR": "Travaux d\u2019accessibilit\u00e9 et d\u2019adaptation pour personnes \u00e0 mobilit\u00e9 r\u00e9duite",
        "INDIS_TEE": "Travaux d\u2019\u00e9conomie d\u2019\u00e9nergie et travaux indissociablement li\u00e9s dans le cadre d\u2019une r\u00e9novation globale",
        "INDIS_PMR": "Travaux d\u2019accessibilit\u00e9 PMR et travaux indissociablement li\u00e9s dans le cadre d\u2019une r\u00e9novation globale",
    }
    type_label_long = type_labels.get(work_type, type_labels["TEE"])
    nb_ops = synthesis["nb_operations"]
    avis_str = ", ".join(str(a).strip() for a in synthesis["num_avis"][:1]) if synthesis["num_avis"] else "N/A"
    tfpb_year = "2024"
    annee_travaux = "2023"

    story.append(Paragraph(
        f'<u>Objet</u>\u00a0: <b>R\u00e9clamation contentieuse en d\u00e9gr\u00e8vement de taxe fonci\u00e8re '
        f'sur les propri\u00e9t\u00e9s b\u00e2ties - Article {ARTICLE_CGI} du '
        f'CGI - {type_label_long} - '
        f'{nb_ops} programmes immobiliers situ\u00e9s \u00e0 {commune.upper()} '
        f'\u2013 Cotisation {tfpb_year} de l\u2019avis n\u00b0{avis_str}</b>',
        styles["MelkoObjet"]
    ))
    story.append(Spacer(1, 2 * mm))

    story.append(Paragraph(
        "<u>Pi\u00e8ces jointes annex\u00e9es</u>\u00a0: "
        "<i>Pi\u00e8ces justificatives et copie de l\u2019avis d\u2019imposition</i>",
        styles["MelkoObjet"]
    ))
    story.append(Spacer(1, 6 * mm))

    # ── Salutation ──
    destinataire_nom = _get_destinataire_nom(synthesis)
    story.append(Paragraph(f"{destinataire_nom},", styles["MelkoBody"]))
    story.append(Spacer(1, 5 * mm))

    # ── Intro paragraph ──
    story.append(Paragraph(
        f'Je soussign\u00e9, mandataire d\u00fbment habilit\u00e9 de {MELKO_INFO["mandant_complet"]}, '
        f'organisme de logement social vis\u00e9 \u00e0 l\u2019article L.411-2 du Code de la Construction '
        f'et de l\u2019Habitation agissant au nom et pour le compte de mon mandant, vous pr\u00e9sente '
        f'une r\u00e9clamation contentieuse en vue de l\u2019obtention du d\u00e9gr\u00e8vement pr\u00e9vu '
        f'\u00e0 l\u2019article {ARTICLE_CGI} du Code G\u00e9n\u00e9ral des Imp\u00f4ts (CGI).',
        styles["MelkoBody"]
    ))

    # ══════════════════════════════════════════════
    # I. PATRIMOINE CONCERNE
    # ══════════════════════════════════════════════
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(
        "I.\u00a0\u00a0\u00a0\u00a0<u>PATRIMOINE CONCERN\u00c9 ET NATURE DES OP\u00c9RATIONS\u00a0:</u>",
        styles["MelkoH1"]
    ))
    story.append(Spacer(1, 2 * mm))

    is_indis = work_type.startswith("INDIS_")

    if is_indis:
        story.append(Paragraph(
            f'La SIP D\u2019HLM a engag\u00e9 en {annee_travaux} un programme de r\u00e9novation '
            f'portant sur son patrimoine locatif social situ\u00e9 \u00e0 {commune}. '
            f'<b>Les d\u00e9penses pr\u00e9sent\u00e9es ci-apr\u00e8s correspondent aux prestations '
            f'indissociablement li\u00e9es</b> aux travaux principaux de ce programme '
            f'(\u00e9tudes pr\u00e9alables, diagnostics, ma\u00eetrise d\u2019\u0153uvre, '
            f'travaux induits, d\u00e9construction).',
            styles["MelkoBody"]
        ))
    else:
        story.append(Paragraph(
            f'La SIP D\u2019HLM a engag\u00e9 en {annee_travaux} un programme ambitieux de '
            f'r\u00e9novation \u00e9nerg\u00e9tique portant sur plusieurs ensembles de son '
            f'patrimoine locatif social situ\u00e9 \u00e0 {commune}.',
            styles["MelkoBody"]
        ))
    story.append(Spacer(1, 2 * mm))

    story.append(Paragraph(
        f'<b>Ce programme global concerne pr\u00e9cis\u00e9ment {nb_ops} op\u00e9rations '
        f'immobili\u00e8res distinctes</b>, chacune r\u00e9pondant \u00e0 des enjeux sp\u00e9cifiques '
        f'de performance \u00e9nerg\u00e9tique et de mise aux normes, dont la liste exhaustive '
        f'est la suivante\u00a0:',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    # ── Operations detail ──
    for i, (op_id, op_info) in enumerate(synthesis["operations"].items(), 1):
        adresses_str = " - ".join(op_info["adresses"][:3]) if op_info["adresses"] else "Adresse non sp\u00e9cifi\u00e9e"
        nb_log = op_info.get("nb_logements", "")
        nb_log_str = f" des {nb_log} logements" if nb_log else ""

        story.append(Paragraph(
            f'<b>{i}.\u00a0\u00a0Op\u00e9ration {op_id} \u2013 {adresses_str}\u00a0:</b>',
            styles["MelkoH2"]
        ))

        if op_info["montant_ht_etudes"] > 0:
            story.append(Spacer(1, 2 * mm))
            story.append(Paragraph(
                f'<i>a.\u00a0\u00a0Prestations d\u2019\u00e9tudes pr\u00e9alables de suivi ou '
                f'd\u2019expertise pour le march\u00e9 de r\u00e9habilitation{nb_log_str}\u00a0:</i>',
                styles["MelkoBullet"]
            ))
            story.append(Paragraph(
                f'\u2013\u00a0\u00a0<b><i>Montant total HT\u00a0: {_format_money(op_info["montant_ht_etudes"])}</i></b>',
                styles["MelkoSubBullet"]
            ))

        if op_info["montant_ht_travaux"] > 0:
            sub_letter = "b" if op_info["montant_ht_etudes"] > 0 else "a"
            story.append(Spacer(1, 2 * mm))
            story.append(Paragraph(
                f'<i>{sub_letter}.\u00a0\u00a0Travaux de r\u00e9novation \u00e9nerg\u00e9tique '
                f'et travaux induits dans le cadre du march\u00e9 de r\u00e9habilitation{nb_log_str}\u00a0:</i>',
                styles["MelkoBullet"]
            ))
            story.append(Paragraph(
                f'\u2013\u00a0\u00a0<b><i>Montant total HT\u00a0: {_format_money(op_info["montant_ht_travaux"])}</i></b>',
                styles["MelkoSubBullet"]
            ))

        story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'<b>Ces {nb_ops} programmes, qui repr\u00e9sentent un investissement significatif dans '
        f'la transition \u00e9nerg\u00e9tique du parc social de la SIP D\u2019HLM \u00e0 {commune}</b>, '
        f'font l\u2019objet de la pr\u00e9sente r\u00e9clamation.',
        styles["MelkoBody"]
    ))

    # ══════════════════════════════════════════════
    # II. FONDEMENT DE LA DEMANDE
    # ══════════════════════════════════════════════
    story.append(PageBreak())
    story.append(Paragraph(
        "II.\u00a0\u00a0\u00a0\u00a0<u>FONDEMENT DE LA DEMANDE ET PRINCIPE D\u2019INDISSOCIABILIT\u00c9\u00a0:</u>",
        styles["MelkoH1"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'<b>Conform\u00e9ment \u00e0 l\u2019article {ARTICLE_CGI} du CGI</b>, nous sollicitons '
        f'le b\u00e9n\u00e9fice du d\u00e9gr\u00e8vement \u00e9gal au quart des d\u00e9penses de '
        f'travaux de r\u00e9novation <b>ayant pour objet de concourir directement \u00e0 la '
        f'r\u00e9alisation d\u2019\u00e9conomies d\u2019\u00e9nergie et de fluides</b>, pay\u00e9es '
        f'au cours de l\u2019ann\u00e9e {annee_travaux} pour l\u2019imposition {tfpb_year}.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'<b>\u00c0 cet \u00e9gard, il convient de rappeler</b> que le champ des travaux \u00e9ligibles '
        f'au d\u00e9gr\u00e8vement est d\u00e9fini en r\u00e9f\u00e9rence \u00e0 ceux ouvrant droit '
        f'au taux r\u00e9duit de TVA de 5,5% en application du 1\u00b0 du 1 du IV de l\u2019article '
        f'278 sexies du CGI. <b>En effet, le BOI-IF-TFB-50-20-20-30 pr\u00e9cise</b> que sont prises '
        f'en compte les d\u00e9penses \u00e9ligibles au taux r\u00e9duit de TVA, pay\u00e9es au cours '
        f'de l\u2019ann\u00e9e pr\u00e9c\u00e9dente.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    # BOI-TVA-IMM reference
    story.append(Paragraph(
        '<b>Or, s\u2019agissant de la d\u00e9finition des travaux \u00e9ligibles</b>, le '
        'BOI-TVA-IMM-20-10-20-10 apporte les pr\u00e9cisions n\u00e9cessaires.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    # §73 citation
    story.append(Paragraph(
        '<b>Ainsi, au \u00a773, il est indiqu\u00e9 que</b> \u00ab\u00a0<i>sont soumis au taux '
        'r\u00e9duit de 5,5% les travaux de r\u00e9novation suivants\u00a0:</i>',
        styles["MelkoBody"]
    ))
    for item in [
        "les travaux d\u2019installation et de remplacement des \u00e9quipements ou de leurs composants\u00a0;",
        "la prestation de main d\u2019\u0153uvre concourant \u00e0 la r\u00e9alisation des travaux de r\u00e9novation\u00a0;",
        "la fourniture d\u2019\u00e9quipement n\u00e9cessaire \u00e0 la r\u00e9novation\u00a0;",
        "les prestations d\u2019\u00e9tudes pr\u00e9alables, de suivi ou d\u2019expertise (notamment les diagnostics pr\u00e9alables aux travaux, les prestations de contr\u00f4le de la conformit\u00e9 des travaux, etc.).\u00a0\u00bb",
    ]:
        story.append(Paragraph(f'-\u00a0\u00a0<i>{item}</i>', styles["MelkoQuoteItem"]))
    story.append(Spacer(1, 3 * mm))

    # §74 reference
    story.append(Paragraph(
        '<b>Par ailleurs, au \u00a774 du m\u00eame BOI</b>, sont express\u00e9ment vis\u00e9s les '
        'travaux concourant directement aux \u00e9conomies d\u2019\u00e9nergie, tels que '
        '<b>les \u00e9l\u00e9ments constitutifs de l\u2019enveloppe du b\u00e2timent, les syst\u00e8mes '
        'de chauffage et d\u2019eau chaude sanitaire, les syst\u00e8mes de ventilation, les syst\u00e8mes '
        'd\u2019\u00e9clairage performants</b> ou encore <b>les \u00e9quipements utilisant des '
        '\u00e9nergies renouvelables</b>.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    # TVA breakdown
    tva_55_ht = synthesis["tva_groups"].get(0.055, {}).get("montant_ht", 0)
    tva_20_ht = synthesis["tva_groups"].get(0.2, {}).get("montant_ht", 0)

    story.append(Paragraph(
        f'<b>En l\u2019esp\u00e8ce, les d\u00e9penses expos\u00e9es</b> s\u2019inscrivent '
        f'parfaitement dans ce cadre l\u00e9gal et doctrinal. '
        f'{"<b>D\u2019une part,</b> elles comprennent d\u2019isolation thermique et de couverture, directement vis\u00e9s au \u00a774 pour un montant de <b>" + _format_money(tva_55_ht) + "</b> au taux de TVA r\u00e9duit de 5,5%." if tva_55_ht > 0 else ""}',
        styles["MelkoBody"]
    ))

    if tva_20_ht > 0:
        story.append(Paragraph(
            f'<b>D\u2019autre part,</b> elles incluent les prestations indissociables de diagnostics '
            f'pr\u00e9alables, d\u2019\u00e9tudes, de ma\u00eetrise d\u2019\u0153uvre, de contr\u00f4le '
            f'technique et d\u2019OPC, express\u00e9ment mentionn\u00e9es au \u00a773 comme \u00e9ligibles, '
            f'pour un montant de <b>{_format_money(tva_20_ht)}</b> au taux de TVA normal.',
            styles["MelkoBody"]
        ))
    story.append(Spacer(1, 3 * mm))

    # Jurisprudence
    story.append(Paragraph(
        '<b>Cette approche inclusive est confirm\u00e9e par la jurisprudence administrative</b>, '
        'laquelle a reconnu \u00e0 plusieurs reprises le caract\u00e8re \u00e9ligible des '
        'd\u00e9penses pr\u00e9paratoires ou induites.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        '<b>En particulier, l\u2019arr\u00eat du Conseil d\u2019\u00c9tat du 17 juin 2015 '
        '(n\u00b0382248)</b> a jug\u00e9 que les frais de ma\u00eetrise d\u2019\u0153uvre et '
        'd\u2019\u00e9tudes pr\u00e9alables sont d\u00e9ductibles d\u00e8s lors qu\u2019ils se '
        'r\u00e9v\u00e8lent n\u00e9cessaires \u00e0 la bonne ex\u00e9cution des travaux principaux.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        '<b>De m\u00eame, les arr\u00eats du 2 juillet 2014 (n\u00b0368070) et du 23 octobre 2015 '
        '(n\u00b0381916)</b> ont confirm\u00e9 la prise en compte des paiements partiels et des '
        'acomptes dans l\u2019assiette du d\u00e9gr\u00e8vement.',
        styles["MelkoBody"]
    ))

    # ══════════════════════════════════════════════
    # III. CALCUL DU DEGREVEMENT
    # ══════════════════════════════════════════════
    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph(
        "III.\u00a0\u00a0\u00a0\u00a0<u>CALCUL DU D\u00c9GR\u00c8VEMENT SOLICIT\u00c9 ET DISPOSITIONS PROC\u00c9DURALES\u00a0:</u>",
        styles["MelkoH1"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'<b>Le montant total des d\u00e9penses \u00e9ligibles</b> pay\u00e9es en {annee_travaux} '
        f's\u2019\u00e9l\u00e8ve \u00e0 <b>{_format_money(synthesis["total_ht_eligible"])}</b>. '
        f'<b>Conform\u00e9ment aux dispositions du BOI-IF-TFB-50-20-20-30 (\u00a770),</b> '
        f'{"aucune subvention n\u2019a \u00e9t\u00e9 per\u00e7ue en " + annee_travaux if synthesis["total_subventions"] == 0 else "les subventions per\u00e7ues s\u2019\u00e9l\u00e8vent \u00e0 " + _format_money(synthesis["total_subventions"])}'
        f', conduisant \u00e0 une base nette de d\u00e9gr\u00e8vement de '
        f'<b>{_format_money(synthesis["base_nette"])}</b>.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'<b>L\u2019application du taux de 25%</b> pr\u00e9vu \u00e0 l\u2019article {ARTICLE_CGI} '
        f'du CGI permet donc de solliciter un d\u00e9gr\u00e8vement total de '
        f'<b>{_format_money(synthesis["total_degrevement"])}</b>, imputable sur les cotisations des '
        f'<b>{nb_ops} programmes concern\u00e9s</b> selon la r\u00e9partition jointe en annexe.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'<b>Au regard des dispositions proc\u00e9durales,</b> la pr\u00e9sente r\u00e9clamation est '
        f'interjet\u00e9e dans le d\u00e9lai de deux ans pr\u00e9vu \u00e0 l\u2019article R*196-2 '
        f'du Livre des Proc\u00e9dures Fiscales, les paiements \u00e9tant intervenus au cours de '
        f'l\u2019ann\u00e9e {annee_travaux}. <b>Conform\u00e9ment \u00e0 l\u2019article R.197-1 '
        f'du LPF</b>, nous vous prions de bien vouloir nous notifier votre d\u00e9cision dans un '
        f'd\u00e9lai de six mois.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    # Pieces jointes
    story.append(Paragraph(
        'En vue d\u2019instruire notre demande, nous vous transmettons un dossier comprenant\u00a0:',
        styles["MelkoBody"]
    ))
    pj_style = ParagraphStyle(
        "MelkoPJ", fontName=FONT_BI, fontSize=10,
        leading=13, leftIndent=30, spaceAfter=3,
    )
    for pj in [
        "L\u2019annexe normalis\u00e9e r\u00e9capitulant les d\u00e9penses retenues pour le calcul du d\u00e9gr\u00e8vement et le d\u00e9tail des travaux consid\u00e9r\u00e9s\u00a0;",
        "Une copie des factures\u00a0;",
        "Une copie des avis de virement valant preuve de paiement\u00a0;",
        "L\u2019avis d\u2019imposition Taxe Fonci\u00e8re {}.".format(tfpb_year),
    ]:
        story.append(Paragraph(f'\u2756\u00a0\u00a0<i>{pj}</i>', pj_style))

    story.append(Spacer(1, 5 * mm))

    # ── Closing ──
    story.append(Paragraph(
        f'<b>Dans ces conditions, et au vu des \u00e9l\u00e9ments fournis,</b> nous vous demandons '
        f'de bien vouloir accorder le d\u00e9gr\u00e8vement de la somme de '
        f'<b>{_format_money(synthesis["total_degrevement"])}</b> \u00e0 '
        f'<b>{MELKO_INFO["mandant_hlm"]}</b> au titre de la cotisation {tfpb_year} de taxe fonci\u00e8re '
        f'sur les propri\u00e9t\u00e9s b\u00e2ties.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        'Nous restons \u00e0 votre disposition pour vous apporter tout compl\u00e9ment '
        'd\u2019informations n\u00e9cessaires \u00e0 l\u2019instruction de notre r\u00e9clamation.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 2 * mm))

    story.append(Paragraph(
        'Par n\u00e9cessit\u00e9 de suivi interne de notre demande, nous vous remercions d\u2019avance '
        'de bien vouloir nous faire suivre par retour de courrier le num\u00e9ro d\u2019affaire '
        'que vous attribuez \u00e0 cette demande.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph(
        f'Nous vous remercions par avance pour l\u2019attention que vous porterez \u00e0 notre demande '
        f'et nous vous prions de croire, {destinataire_nom}, en l\u2019expression de nos sinc\u00e8res '
        f'salutations.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 8 * mm))

    # ── Signature ──
    sig_style_blue = ParagraphStyle(
        "SigBlue", fontName=FONT_B, fontSize=11,
        textColor=BLUE_ACCENT, leading=14,
    )
    sig_style_title = ParagraphStyle(
        "SigTitle", fontName=FONT_B, fontSize=10,
        textColor=BLUE_DARK, leading=13,
    )
    sig_style_company = ParagraphStyle(
        "SigCompany", fontName=FONT_B, fontSize=10,
        leading=13,
    )
    story.append(Paragraph(MELKO_INFO["signataire_nom"], sig_style_blue))
    story.append(Paragraph(MELKO_INFO["signataire_titre"], sig_style_title))
    story.append(Paragraph(MELKO_INFO["nom"], sig_style_company))

    if LOGO_PATH.exists():
        story.append(Spacer(1, 5 * mm))
        logo_small = Image(str(LOGO_PATH), width=22 * mm, height=22 * mm)
        logo_small.hAlign = "LEFT"
        story.append(logo_small)

    # ── Build ──
    doc.build(story, onFirstPage=_add_footer, onLaterPages=_add_later_pages)

    log.info(f"Courrier PDF genere : {file_path}")
    return file_path


# ══════════════════════════════════════════════
# Helper functions
# ══════════════════════════════════════════════

def _add_header_block(story, styles, synthesis, commune):
    """Add the sender/recipient two-column header."""
    sender_style = ParagraphStyle(
        "sender", fontName=FONT, fontSize=10, leading=13,
    )
    recipient_style = ParagraphStyle(
        "recipient", fontName=FONT, fontSize=10, leading=13, alignment=TA_RIGHT,
    )

    sender_text = (
        f'{MELKO_INFO["nom"]}<br/>'
        f'Pour le compte de {MELKO_INFO["mandant"]}<br/>'
        f'{MELKO_INFO["adresse_mandant_ligne1"]}<br/>'
        f'{MELKO_INFO["adresse_mandant_ville"]}'
    )

    cdif_name = synthesis["cdif"][0] if synthesis["cdif"] else "SDIF"
    adresse_cdif = synthesis["adresse_cdif"][0] if synthesis["adresse_cdif"] else ""
    recipient_lines = adresse_cdif.replace("\n", "<br/>") if adresse_cdif else cdif_name

    sender_para = Paragraph(sender_text, sender_style)
    recipient_para = Paragraph(recipient_lines, recipient_style)

    header_table = Table(
        [[sender_para, recipient_para]],
        colWidths=[85 * mm, 85 * mm],
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)


def _add_footer(canvas, doc):
    """Footer on the first page (no top band)."""
    _draw_footer(canvas)


def _add_later_pages(canvas, doc):
    """Footer + blue top band on pages 2+."""
    canvas.saveState()
    w, h = A4
    # Blue gradient band at top
    band_height = 8 * mm
    canvas.setFillColor(BLUE_ACCENT)
    canvas.rect(0, h - band_height, w, band_height, stroke=0, fill=1)
    canvas.restoreState()
    _draw_footer(canvas)


def _draw_footer(canvas):
    """Draw the blue footer with company info."""
    canvas.saveState()
    w, _ = A4

    canvas.setStrokeColor(BLUE_ACCENT)
    canvas.setLineWidth(1)
    canvas.line(20 * mm, 15 * mm, w - 20 * mm, 15 * mm)

    canvas.setFont(FONT_B if _cal_ok else "Helvetica-Bold", 6.5)
    canvas.setFillColor(BLUE_ACCENT)
    footer_text = (
        f'{MELKO_INFO["adresse_ligne1"]} {MELKO_INFO["adresse_ligne2"]} - '
        f'SIREN\u00a0: {MELKO_INFO["siren"]} - {MELKO_INFO["rcs"]} - '
        f'{MELKO_INFO["forme"]} AU CAPITAL DE {MELKO_INFO["capital"]}\u20ac'
    )
    footer_text2 = (
        f'SIRET\u00a0: {MELKO_INFO["siret"]} - '
        f'N\u00b0 TVA INTRACOMMUNAUTAIRE\u00a0: {MELKO_INFO["tva_intra"]}'
    )
    canvas.drawCentredString(w / 2, 11 * mm, footer_text)
    canvas.drawCentredString(w / 2, 7.5 * mm, footer_text2)

    canvas.restoreState()


def _get_destinataire_nom(synthesis) -> str:
    """Extract the name of the recipient from CDIF address."""
    for addr in synthesis.get("adresse_cdif", []):
        addr_str = str(addr)
        if "attention de" in addr_str.lower():
            parts = addr_str.lower().split("attention de")
            if len(parts) > 1:
                name = parts[1].strip().split("\n")[0].strip()
                # Remove leading "monsieur"/"madame" if already present
                name_upper = name.upper()
                for prefix in ("MONSIEUR ", "MADAME ", "M. ", "MME "):
                    if name_upper.startswith(prefix):
                        return f"{prefix.strip().title()} {name_upper[len(prefix):]}"
                return f"Monsieur {name_upper}"
    return "Madame, Monsieur"
