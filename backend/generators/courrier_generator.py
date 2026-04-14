from pathlib import Path
from datetime import datetime
import locale

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle,
    PageBreak, KeepTogether,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.platypus.flowables import HRFlowable

from backend.config.constants import MELKO_INFO, LOGO_PATH, ARTICLE_CGI
from backend.utils.logger import get_logger

log = get_logger()

# ── Colors ──
BLUE_DARK = HexColor("#1F4E79")
BLUE_ACCENT = HexColor("#2E86C1")
BLUE_LIGHT = HexColor("#D6EAF8")
GREY_FOOTER = HexColor("#888888")


def _get_styles():
    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(
        "MelkoTitle", fontName="Helvetica-Bold", fontSize=10,
        textColor=BLUE_DARK, leading=14,
    ))
    styles.add(ParagraphStyle(
        "MelkoBody", fontName="Helvetica", fontSize=9,
        leading=13, alignment=TA_JUSTIFY, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "MelkoBodyBold", fontName="Helvetica-Bold", fontSize=9,
        leading=13, alignment=TA_JUSTIFY, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "MelkoH1", fontName="Helvetica-Bold", fontSize=10,
        textColor=BLUE_DARK, leading=14, spaceAfter=8, spaceBefore=12,
    ))
    styles.add(ParagraphStyle(
        "MelkoH2", fontName="Helvetica-Bold", fontSize=9,
        leading=12, spaceAfter=6, spaceBefore=8,
    ))
    styles.add(ParagraphStyle(
        "MelkoSmall", fontName="Helvetica", fontSize=7,
        textColor=GREY_FOOTER, leading=9, alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "MelkoObjet", fontName="Helvetica-Bold", fontSize=9,
        textColor=BLUE_ACCENT, leading=12, spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "MelkoBullet", fontName="Helvetica", fontSize=9,
        leading=12, leftIndent=20, spaceAfter=3,
    ))
    return styles


def _format_money(amount: float) -> str:
    """Format a number as French currency."""
    if amount == 0:
        return "0 €"
    formatted = f"{amount:,.2f}".replace(",", " ").replace(".", ",")
    return f"{formatted} €"


def _format_date_fr() -> str:
    days = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
    months = ["janvier", "février", "mars", "avril", "mai", "juin",
              "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    now = datetime.now()
    return f"le {days[now.weekday()]} {now.day} {months[now.month-1]} {now.year}"


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
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=12*mm, bottomMargin=18*mm,
    )

    styles = _get_styles()
    story = []

    # ── Logo + Sender/Recipient on same level ──
    if LOGO_PATH.exists():
        logo = Image(str(LOGO_PATH), width=22*mm, height=22*mm)
        logo.hAlign = "LEFT"
        story.append(logo)
    story.append(Spacer(1, 2*mm))

    # ── Sender / Recipient block ──
    _add_header_block(story, styles, synthesis, commune)

    # ── Date ──
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph(f"Fait à Paris, {_format_date_fr()},", styles["MelkoBody"]))
    story.append(Spacer(1, 2*mm))

    # ── Affaire suivie par ──
    story.append(Paragraph(
        f'<u>Affaire suivie par</u> : {MELKO_INFO["signataire_nom"]}, '
        f'{MELKO_INFO["signataire_email"]} (demande N°{demande_num})',
        styles["MelkoObjet"]
    ))
    story.append(Spacer(1, 2*mm))

    # ── Objet ──
    type_labels = {
        "TEE": "Travaux d'économie d'énergie dans le cadre d'une rénovation globale",
        "PMR": "Travaux d'accessibilité et d'adaptation pour personnes à mobilité réduite",
        "INDIS_TEE": "Travaux indissociablement liés aux travaux d'économie d'énergie (études, diagnostics, maîtrise d'œuvre, travaux induits)",
        "INDIS_PMR": "Travaux indissociablement liés aux travaux d'accessibilité PMR (études, diagnostics, maîtrise d'œuvre, travaux induits)",
    }
    type_label_long = type_labels.get(work_type, type_labels["TEE"])
    nb_ops = synthesis["nb_operations"]
    code_postal = ""
    avis_str = ", ".join(str(a).strip() for a in synthesis["num_avis"][:1]) if synthesis["num_avis"] else "N/A"
    tfpb_year = "2024"

    story.append(Paragraph(
        f'<u>Objet</u> : <b>Réclamation contentieuse en dégrèvement de taxe foncière sur les propriétés bâties '
        f'- Article {ARTICLE_CGI} du CGI - {type_label_long} - '
        f'{nb_ops} programmes immobiliers situés à {commune.upper()} '
        f'– Cotisation {tfpb_year} de l\'avis n°{avis_str}</b>',
        styles["MelkoObjet"]
    ))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph(
        "<u>Pièces jointes annexées</u> : <i>Pièces justificatives et copie de l'avis d'imposition</i>",
        styles["MelkoObjet"]
    ))
    story.append(Spacer(1, 3*mm))

    # ── Salutation ──
    destinataire_nom = _get_destinataire_nom(synthesis)
    story.append(Paragraph(f"{destinataire_nom},", styles["MelkoBody"]))
    story.append(Spacer(1, 3*mm))

    # ── Intro paragraph ──
    story.append(Paragraph(
        f'Je soussigné, mandataire dûment habilité de {MELKO_INFO["mandant_complet"]}, '
        f'organisme de logement social visé à l\'article L.411-2 du Code de la Construction et de l\'Habitation '
        f'agissant au nom et pour le compte de mon mandant, vous présente une réclamation contentieuse en vue '
        f'de l\'obtention du dégrèvement prévu à l\'article {ARTICLE_CGI} du Code Général des Impôts (CGI).',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3*mm))

    # ── I. PATRIMOINE CONCERNÉ ──
    story.append(Paragraph("I. PATRIMOINE CONCERNÉ ET NATURE DES OPÉRATIONS :", styles["MelkoH1"]))
    story.append(Spacer(1, 2*mm))

    annee_travaux = "2023"
    is_indis = work_type.startswith("INDIS_")

    if is_indis:
        story.append(Paragraph(
            f'La SIP D\'HLM a engagé en {annee_travaux} un programme de rénovation '
            f'portant sur son patrimoine locatif social situé à {commune}. '
            f'<b>Les dépenses présentées ci-après correspondent aux prestations indissociablement liées</b> '
            f'aux travaux principaux de ce programme (études préalables, diagnostics, maîtrise d\'œuvre, '
            f'travaux induits, déconstruction).',
            styles["MelkoBody"]
        ))
    else:
        story.append(Paragraph(
            f'La SIP D\'HLM a engagé en {annee_travaux} un programme ambitieux de rénovation énergétique '
            f'portant sur plusieurs ensembles de son patrimoine locatif social situé à {commune}.',
            styles["MelkoBody"]
        ))

    story.append(Paragraph(
        f'<b>Ce programme global concerne précisément {nb_ops} opérations immobilières distinctes</b>, '
        f'chacune répondant à des enjeux spécifiques de performance énergétique et de mise aux normes, '
        f'dont la liste exhaustive est la suivante :',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3*mm))

    # ── Operations detail ──
    for i, (op_id, op_info) in enumerate(synthesis["operations"].items(), 1):
        adresses_str = " - ".join(op_info["adresses"][:3]) if op_info["adresses"] else "Adresse non spécifiée"
        nb_log = op_info.get("nb_logements", "")
        nb_log_str = f" des {nb_log} logements" if nb_log else ""

        story.append(Paragraph(
            f'<b>{i}. Opération {op_id} – {adresses_str} :</b>',
            styles["MelkoH2"]
        ))

        if op_info["montant_ht_etudes"] > 0:
            story.append(Paragraph(
                f'a. Prestations d\'études préalables de suivi ou d\'expertise pour le marché de réhabilitation{nb_log_str} :',
                styles["MelkoBullet"]
            ))
            story.append(Paragraph(
                f'– Montant total HT : <b>{_format_money(op_info["montant_ht_etudes"])}</b>',
                styles["MelkoBullet"]
            ))

        if op_info["montant_ht_travaux"] > 0:
            story.append(Paragraph(
                f'b. Travaux de rénovation énergétique et travaux induits{nb_log_str} :',
                styles["MelkoBullet"]
            ))
            story.append(Paragraph(
                f'– Montant total HT : <b>{_format_money(op_info["montant_ht_travaux"])}</b>',
                styles["MelkoBullet"]
            ))

        story.append(Spacer(1, 3*mm))

    story.append(Paragraph(
        f'Ces {nb_ops} programmes, qui représentent un investissement significatif dans la transition '
        f'énergétique du parc social de la SIP D\'HLM à {commune}, font l\'objet de la présente réclamation.',
        styles["MelkoBody"]
    ))

    # ── II. FONDEMENT ──
    story.append(PageBreak())
    story.append(Paragraph("II. FONDEMENT DE LA DEMANDE ET PRINCIPE D'INDISSOCIABILITÉ :", styles["MelkoH1"]))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph(
        f'<b>Conformément à l\'article {ARTICLE_CGI} du CGI</b>, nous sollicitons le bénéfice du dégrèvement '
        f'égal au quart des dépenses de travaux de rénovation <b>ayant pour objet de concourir directement '
        f'à la réalisation d\'économies d\'énergie et de fluides</b>, payées au cours de l\'année {annee_travaux} '
        f'pour l\'imposition {tfpb_year}.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 2*mm))

    if is_indis:
        story.append(Paragraph(
            '<b>Les dépenses présentées dans le présent dossier correspondent aux prestations '
            'indissociablement liées aux travaux principaux</b>, au sens du BOI-IF-TFB-50-20-20-30 (§74 et suivants). '
            'Il s\'agit notamment des prestations d\'études préalables, de diagnostics, de maîtrise d\'œuvre, '
            'de contrôle technique, d\'OPC, ainsi que des travaux induits par la rénovation (dépose, '
            'déconstruction, adaptation des réseaux).',
            styles["MelkoBody"]
        ))
        story.append(Spacer(1, 2*mm))
        story.append(Paragraph(
            '<b>Le caractère indissociable de ces prestations</b> résulte du fait qu\'elles constituent '
            'une condition sine qua non de la réalisation des travaux principaux d\'économie d\'énergie. '
            'Sans ces prestations préalables ou induites, les travaux de rénovation énergétique ne '
            'pourraient être menés à bien.',
            styles["MelkoBody"]
        ))
        story.append(Spacer(1, 2*mm))

    story.append(Paragraph(
        '<b>À cet égard, il convient de rappeler</b> que le champ des travaux éligibles au dégrèvement est défini '
        'en référence à ceux ouvrant droit au taux réduit de TVA de 5,5% en application du 1° du 1 du IV de '
        'l\'article 278 sexies du CGI. <b>En effet, le BOI-IF-TFB-50-20-20-30 précise</b> que sont prises en compte '
        'les dépenses éligibles au taux réduit de TVA, payées au cours de l\'année précédente.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 2*mm))

    # TVA breakdown
    tva_55_ht = synthesis["tva_groups"].get(0.055, {}).get("montant_ht", 0)
    tva_20_ht = synthesis["tva_groups"].get(0.2, {}).get("montant_ht", 0)

    if tva_55_ht > 0:
        story.append(Paragraph(
            f'En l\'espèce, les dépenses exposées comprennent d\'isolation thermique et de couverture, '
            f'directement visés au §74 pour un montant de <b>{_format_money(tva_55_ht)}</b> au taux de TVA réduit de 5,5%.',
            styles["MelkoBody"]
        ))

    if tva_20_ht > 0:
        story.append(Paragraph(
            f'D\'autre part, elles incluent les prestations indissociables de diagnostics préalables, d\'études, '
            f'de maîtrise d\'œuvre, de contrôle technique et d\'OPC, pour un montant de '
            f'<b>{_format_money(tva_20_ht)}</b> au taux de TVA normal.',
            styles["MelkoBody"]
        ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph(
        '<b>Cette approche inclusive est confirmée par la jurisprudence administrative</b>, laquelle a reconnu '
        'à plusieurs reprises le caractère éligible des dépenses préparatoires ou induites.',
        styles["MelkoBody"]
    ))
    story.append(Paragraph(
        'En particulier, l\'arrêt du <b>Conseil d\'État du 17 juin 2015 (n°382248)</b> a jugé que les frais de '
        'maîtrise d\'œuvre et d\'études préalables sont déductibles.',
        styles["MelkoBody"]
    ))
    story.append(Paragraph(
        'De même, les arrêts du <b>2 juillet 2014 (n°368070) et du 23 octobre 2015 (n°381916)</b> ont confirmé '
        'la prise en compte des paiements partiels et des acomptes.',
        styles["MelkoBody"]
    ))

    # ── III. CALCUL ──
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph("III. CALCUL DU DÉGRÈVEMENT SOLICITÉ ET DISPOSITIONS PROCÉDURALES :", styles["MelkoH1"]))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph(
        f'<b>Le montant total des dépenses éligibles</b> payées en {annee_travaux} s\'élève à '
        f'<b>{_format_money(synthesis["total_ht_eligible"])}</b>. '
        f'Conformément aux dispositions du BOI-IF-TFB-50-20-20-30 (§70), '
        f'{"aucune subvention n\'a été perçue en " + annee_travaux if synthesis["total_subventions"] == 0 else "les subventions perçues s\'élèvent à " + _format_money(synthesis["total_subventions"])}'
        f', conduisant à une base nette de dégrèvement de <b>{_format_money(synthesis["base_nette"])}</b>.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 2*mm))

    story.append(Paragraph(
        f'L\'application du taux de 25% prévu à l\'article {ARTICLE_CGI} du CGI permet donc de solliciter '
        f'un dégrèvement total de <b>{_format_money(synthesis["total_degrevement"])}</b>, imputable sur les '
        f'cotisations des {nb_ops} programmes concernés selon la répartition jointe en annexe.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph(
        'Au regard des dispositions procédurales, la présente réclamation est interjetée dans le délai de '
        'deux ans prévu à l\'article R*196-2 du Livre des Procédures Fiscales. '
        'Conformément à l\'article R.197-1 du LPF, nous vous prions de bien vouloir nous notifier '
        'votre décision dans un délai de six mois.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph('En vue d\'instruire notre demande, nous vous transmettons un dossier comprenant :', styles["MelkoBody"]))
    for pj in [
        "L'annexe normalisée récapitulant les dépenses retenues pour le calcul du dégrèvement ;",
        "Une copie des factures ;",
        "Une copie des avis de virement valant preuve de paiement ;",
        "L'avis d'imposition Taxe Foncière 2024.",
    ]:
        story.append(Paragraph(f"• <i>{pj}</i>", styles["MelkoBullet"]))

    story.append(Spacer(1, 5*mm))

    # ── Closing ──
    story.append(Paragraph(
        f'Dans ces conditions, et au vu des éléments fournis, nous vous demandons de bien vouloir accorder '
        f'le dégrèvement de la somme de <b>{_format_money(synthesis["total_degrevement"])}</b> à '
        f'<b>{MELKO_INFO["mandant_hlm"]}</b> au titre de la cotisation {tfpb_year} de taxe foncière sur les propriétés bâties.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph(
        'Nous restons à votre disposition pour vous apporter tout complément d\'informations nécessaires.',
        styles["MelkoBody"]
    ))
    story.append(Paragraph(
        'Par nécessité de suivi interne, nous vous remercions de bien vouloir nous faire suivre '
        'par retour de courrier le numéro d\'affaire que vous attribuez à cette demande.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 3*mm))

    story.append(Paragraph(
        f'Nous vous prions de croire, {destinataire_nom}, en l\'expression de nos sincères salutations.',
        styles["MelkoBody"]
    ))
    story.append(Spacer(1, 5*mm))

    # ── Signature ──
    story.append(Paragraph(f'<b>{MELKO_INFO["signataire_nom"]}</b>', styles["MelkoBodyBold"]))
    story.append(Paragraph(f'<b>{MELKO_INFO["signataire_titre"]}</b>', styles["MelkoBody"]))
    story.append(Paragraph(f'<b>{MELKO_INFO["nom"]}</b>', styles["MelkoBody"]))

    if LOGO_PATH.exists():
        story.append(Spacer(1, 3*mm))
        logo_small = Image(str(LOGO_PATH), width=18*mm, height=18*mm)
        logo_small.hAlign = "LEFT"
        story.append(logo_small)

    # ── Build with footer ──
    doc.build(story, onFirstPage=_add_footer, onLaterPages=_add_footer)

    log.info(f"Courrier PDF genere : {file_path}")
    return file_path


def _add_header_block(story, styles, synthesis, commune):
    """Add the sender/recipient two-column header."""
    # Sender (left)
    sender_text = (
        f'{MELKO_INFO["nom"]}<br/>'
        f'Pour le compte de {MELKO_INFO["mandant"]}<br/>'
        f'{MELKO_INFO["adresse_mandant_ligne1"]}<br/>'
        f'{MELKO_INFO["adresse_mandant_ville"]}'
    )

    # Recipient (right)
    cdif_name = synthesis["cdif"][0] if synthesis["cdif"] else "SDIF"
    adresse_cdif = synthesis["adresse_cdif"][0] if synthesis["adresse_cdif"] else ""

    recipient_lines = adresse_cdif.replace("\n", "<br/>") if adresse_cdif else cdif_name

    sender_para = Paragraph(sender_text, ParagraphStyle(
        "sender", fontName="Helvetica", fontSize=9, leading=12,
    ))
    recipient_para = Paragraph(recipient_lines, ParagraphStyle(
        "recipient", fontName="Helvetica", fontSize=9, leading=12, alignment=TA_RIGHT,
    ))

    header_table = Table(
        [[sender_para, recipient_para]],
        colWidths=[85*mm, 85*mm],
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)


def _add_footer(canvas, doc):
    """Add the blue footer with company info on every page."""
    canvas.saveState()
    w, h = A4

    # Blue line
    canvas.setStrokeColor(BLUE_ACCENT)
    canvas.setLineWidth(1)
    canvas.line(20*mm, 15*mm, w - 20*mm, 15*mm)

    # Footer text
    canvas.setFont("Helvetica", 6.5)
    canvas.setFillColor(GREY_FOOTER)
    footer_text = (
        f'{MELKO_INFO["adresse_ligne1"]} {MELKO_INFO["adresse_ligne2"]} - '
        f'SIREN : {MELKO_INFO["siren"]} - {MELKO_INFO["rcs"]} - '
        f'{MELKO_INFO["forme"]} AU CAPITAL DE {MELKO_INFO["capital"]}€'
    )
    footer_text2 = (
        f'SIRET : {MELKO_INFO["siret"]} - '
        f'N° TVA INTRACOMMUNAUTAIRE : {MELKO_INFO["tva_intra"]}'
    )
    canvas.drawCentredString(w/2, 11*mm, footer_text)
    canvas.drawCentredString(w/2, 7.5*mm, footer_text2)

    canvas.restoreState()


def _get_destinataire_nom(synthesis) -> str:
    """Extract the name of the recipient from CDIF address."""
    for addr in synthesis.get("adresse_cdif", []):
        addr_str = str(addr)
        if "attention de" in addr_str.lower():
            parts = addr_str.lower().split("attention de")
            if len(parts) > 1:
                name = parts[1].strip().split("\n")[0].strip()
                return f"Monsieur {name.title()}"
    return "Madame, Monsieur"
