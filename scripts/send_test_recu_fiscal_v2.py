"""
v2 — Envoi d'un reçu fiscal PDF (logo + signature) via API Gmail.
     Email HTML avec logo inline. Vouvoiement.
     PDF généré avec ReportLab (pure Python, pas de dép système).

Prérequis :
    - python scripts/gmail_auth.py (avec scope gmail.send)
    - Gmail API activée côté Google Cloud
    - Fichiers à la racine :
        Logo_Monique_Festival.png
        Signature_Judith_.png
    - pip install reportlab

Usage : python scripts/send_test_recu_fiscal_v2.py
"""
import sys
import os
import json
import base64
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, KeepTogether, HRFlowable
)

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO_ROOT = Path(__file__).resolve().parent.parent

LOGO_PATH = REPO_ROOT / "Logo_Monique_Festival.png"
SIGNATURE_PATH = REPO_ROOT / "Signature_Judith_.png"
OUTPUT_DIR = REPO_ROOT / "officiel/dons/recus_emis"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ======== Association ========
ASSO = {
    "raison_sociale": "MONIQUE FESTIVAL",
    "ancien_nom": "anciennement JD Production",
    "objet": (
        "Organisation annuelle d'un festival pluridisciplinaire à Besançon et en Bourgogne-Franche-Comté, "
        "mêlant musique classique, musiques actuelles, théâtre et arts vivants ; soutien à la création "
        "artistique émergente ; actions d'éducation artistique et culturelle, de médiation et de transmission, "
        "notamment vers les publics éloignés de la pratique ; valorisation du territoire ; promotion de la "
        "diffusion de la culture, à titre non lucratif et d'intérêt général."
    ),
    "rna": "W251011013",
    "siren": "991 055 450",
    "siret": "991 055 450 00019",
    "adresse": "54 chemin de Valentin",
    "cp": "25000",
    "ville": "Besançon",
    "pays": "France",
    "signataire": "Judith LAITHIER",
    "fonction": "Présidente",
    "ville_sig": "Besançon",
}

# ======== Donation test ========
DON = {
    "donateur_prenom": "Charles",
    "donateur_nom": "JANIER-DUBRY",
    "donateur_adresse": "54 chemin de Valentin",
    "donateur_cp": "25000",
    "donateur_ville": "Besançon",
    "donateur_pays": "France",
    "email": "charles@janier-dubry.fr",
    "montant": 50,
    "montant_lettres": "cinquante euros",
    "date_don": "24/04/2026",
    "moyen_paiement": "Carte bancaire",
    "numero_recu": "TEST-2026-04-24-0001",
    "date_emission": "24/04/2026",
    "article_cgi": "200",
    "taux_reduction": "66",
    "is_test": True,
}


def load_creds():
    with open(TOKEN_FILE) as f:
        data = json.load(f)
    if 'https://www.googleapis.com/auth/gmail.send' not in data.get('scopes', []):
        sys.exit("ERREUR : scope gmail.send manquant. Lancez python scripts/gmail_auth.py")
    creds = Credentials(
        token=data['token'],
        refresh_token=data.get('refresh_token'),
        token_uri=data['token_uri'],
        client_id=data['client_id'],
        client_secret=data['client_secret'],
        scopes=data['scopes'],
    )
    if not creds.valid and creds.refresh_token:
        creds.refresh(Request())
        data['token'] = creds.token
        with open(TOKEN_FILE, 'w') as f:
            json.dump(data, f, indent=2)
    return creds


# ============================================================================
# GÉNÉRATION PDF avec ReportLab
# ============================================================================

def build_pdf(don, output_path: Path):
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
        leftMargin=2.2 * cm,
        rightMargin=2.2 * cm,
        title=f"Reçu fiscal {don['numero_recu']} — {ASSO['raison_sociale']}",
        author=ASSO['raison_sociale'],
    )

    styles = getSampleStyleSheet()
    brand = colors.HexColor("#d12d2d")
    dark = colors.HexColor("#111111")
    grey = colors.HexColor("#555555")
    light = colors.HexColor("#f7f7f7")

    story = []

    # ======== EN-TÊTE : logo + identité ========
    logo_img = Image(str(LOGO_PATH), width=2.2 * cm, height=2.2 * cm) if LOGO_PATH.exists() else Paragraph("", styles['Normal'])

    st_title = ParagraphStyle(name="ht", fontSize=18, leading=22, textColor=dark, fontName="Helvetica-Bold")
    st_sub = ParagraphStyle(name="hs", fontSize=8, leading=10, textColor=grey, fontName="Helvetica-Oblique")
    st_line = ParagraphStyle(name="hl", fontSize=8, leading=10, textColor=colors.HexColor("#777777"))

    identite_block = [
        Paragraph(ASSO['raison_sociale'], st_title),
        Paragraph(ASSO['ancien_nom'], st_sub),
        Paragraph("Association Loi 1901 — Œuvre ou organisme d'intérêt général — Arts et culture",
                  ParagraphStyle(name="hss", fontSize=8, leading=10, textColor=grey)),
        Paragraph(f"RNA {ASSO['rna']} · SIREN {ASSO['siren']} · SIRET {ASSO['siret']}", st_line),
        Paragraph(f"{ASSO['adresse']} — {ASSO['cp']} {ASSO['ville']} — {ASSO['pays']}", st_line),
    ]

    header_table = Table(
        [[logo_img, identite_block]],
        colWidths=[2.6 * cm, None]
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1.5, color=dark))
    story.append(Spacer(1, 14))

    # ======== TITRE ========
    story.append(Paragraph(
        '<para alignment="center" fontSize="16"><b>REÇU AU TITRE DES DONS</b></para>',
        styles['Normal']
    ))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        f'<para alignment="center" fontSize="10" textColor="#555555">'
        f'Article {don["article_cgi"]} du Code Général des Impôts</para>',
        styles['Normal']
    ))
    story.append(Spacer(1, 12))

    # ======== Numéro d'ordre ========
    numero_style = ParagraphStyle(
        name="num", alignment=TA_CENTER, fontSize=11, textColor=dark,
        borderColor=dark, borderWidth=1, borderPadding=6, leading=13,
    )
    story.append(Paragraph(f"<b>N° d'ordre : {don['numero_recu']}</b>", numero_style))
    story.append(Spacer(1, 14))

    # ======== Donateur + Bénéficiaire côte à côte ========
    donateur_html = (
        f"<b>{don['donateur_prenom']} {don['donateur_nom']}</b><br/>"
        f"{don['donateur_adresse']}<br/>"
        f"{don['donateur_cp']} {don['donateur_ville']}<br/>"
        f"{don['donateur_pays']}"
    )
    benef_html = (
        f"<b>{ASSO['raison_sociale']}</b><br/>"
        f"{ASSO['adresse']}<br/>"
        f"{ASSO['cp']} {ASSO['ville']}<br/>"
        f"{ASSO['pays']}"
    )
    style_block = ParagraphStyle(name="blk", fontSize=10, leading=14)
    style_block_title = ParagraphStyle(
        name="blktit", fontSize=9, textColor=grey, spaceAfter=4,
        fontName="Helvetica-Bold"
    )

    col_don = [
        Paragraph("DONATEUR", style_block_title),
        Paragraph(donateur_html, style_block),
    ]
    col_benef = [
        Paragraph("BÉNÉFICIAIRE", style_block_title),
        Paragraph(benef_html, style_block),
    ]
    blocks_table = Table(
        [[col_don, col_benef]],
        colWidths=[None, None]
    )
    blocks_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (0, 0), 0),
        ('RIGHTPADDING', (1, 0), (1, 0), 0),
        ('LEFTPADDING', (1, 0), (1, 0), 12),
    ]))
    story.append(blocks_table)
    story.append(Spacer(1, 14))

    # ======== Objet ========
    story.append(Paragraph("OBJET DE L'ASSOCIATION", style_block_title))
    objet_style = ParagraphStyle(
        name="objet", fontSize=9.5, leading=13, leftIndent=10, rightIndent=10,
        borderPadding=10, backColor=light, borderColor=brand, borderWidth=0,
    )
    # Tableau fond gris + bordure gauche rouge (astuce reportlab)
    objet_box = Table([[Paragraph(ASSO['objet'], objet_style)]], colWidths=[None])
    objet_box.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), light),
        ('LINEBEFORE', (0, 0), (0, -1), 3, brand),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    story.append(objet_box)
    story.append(Spacer(1, 14))

    # ======== Montant ========
    montant_style = ParagraphStyle(
        name="mnt", alignment=TA_CENTER, fontSize=14, textColor=dark,
        leading=18,
    )
    montant_sub_style = ParagraphStyle(
        name="mntsub", alignment=TA_CENTER, fontSize=9.5, textColor=grey,
        leading=12,
    )
    montant_content = [
        Paragraph(
            f"<b>{don['montant']} € ({don['montant_lettres']})</b>",
            montant_style
        ),
        Spacer(1, 4),
        Paragraph(
            f"Date du don : <b>{don['date_don']}</b> &nbsp;·&nbsp; "
            f"Forme : Don manuel &nbsp;·&nbsp; "
            f"Nature : Numéraire &nbsp;·&nbsp; "
            f"Moyen de paiement : {don['moyen_paiement']}",
            montant_sub_style
        ),
    ]
    montant_box = Table([[montant_content]], colWidths=[None])
    montant_box.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#fff4f4")),
        ('BOX', (0, 0), (-1, -1), 1, brand),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
    ]))
    story.append(montant_box)
    story.append(Spacer(1, 14))

    # ======== Certification ========
    cert_style = ParagraphStyle(
        name="cert", fontSize=9, textColor="#444444", leading=12,
        fontName="Helvetica-Oblique", leftIndent=4, rightIndent=4,
    )
    cert_text = (
        f"Le bénéficiaire certifie sur l'honneur que les dons et versements qu'il reçoit "
        f"ouvrent droit à la réduction d'impôt prévue à l'article {don['article_cgi']} "
        f"du Code Général des Impôts.<br/><br/>"
        f"<b>Particulier</b> : vous pouvez bénéficier d'une réduction d'impôt égale à "
        f"<b>{don['taux_reduction']} %</b> du montant de votre don, dans la limite de 20 % "
        f"de votre revenu imposable."
    )
    cert_box = Table([[Paragraph(cert_text, cert_style)]], colWidths=[None])
    cert_box.setStyle(TableStyle([
        ('LINEABOVE', (0, 0), (-1, 0), 0.5, colors.HexColor("#cccccc")),
        ('LINEBELOW', (0, -1), (-1, -1), 0.5, colors.HexColor("#cccccc")),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(cert_box)
    story.append(Spacer(1, 22))

    # ======== Lieu/date + signature ========
    lieudate_style = ParagraphStyle(name="ld", fontSize=10, leading=14)
    lieudate_html = (
        f"Fait à <b>{ASSO['ville_sig']}</b>,<br/>"
        f"le <b>{don['date_emission']}</b>"
    )

    if SIGNATURE_PATH.exists():
        sig_img = Image(str(SIGNATURE_PATH), width=3.5 * cm, height=1.5 * cm, kind='proportional')
    else:
        sig_img = Paragraph('<i>(signature manuscrite)</i>', lieudate_style)

    sig_name_style = ParagraphStyle(name="sn", alignment=TA_CENTER, fontSize=10, leading=13, fontName="Helvetica-Bold")
    sig_func_style = ParagraphStyle(name="sf", alignment=TA_CENTER, fontSize=9, textColor=grey, leading=12)

    sig_block = [
        sig_img,
        Paragraph(ASSO['signataire'], sig_name_style),
        Paragraph(f"{ASSO['fonction']} de l'association {ASSO['raison_sociale']}", sig_func_style),
    ]

    sig_table = Table(
        [[Paragraph(lieudate_html, lieudate_style), sig_block]],
        colWidths=[7 * cm, None]
    )
    sig_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
    ]))
    story.append(sig_table)
    story.append(Spacer(1, 18))

    # ======== Footer ========
    footer_style = ParagraphStyle(
        name="ftr", fontSize=8, alignment=TA_CENTER, textColor=colors.HexColor("#888888")
    )
    footer_text = "Reçu émis conformément au modèle Cerfa n° 11580*04. À conserver pour votre déclaration d'impôts."
    if don.get('is_test'):
        footer_text += '<br/><b><font color="#d12d2d">⚠ RECU TEST — À des fins de démonstration — Ne pas utiliser pour déclaration fiscale</font></b>'
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bbbbbb"), dash=(2, 2)))
    story.append(Spacer(1, 6))
    story.append(Paragraph(footer_text, footer_style))

    doc.build(story)
    return output_path


# ============================================================================
# EMAIL HTML
# ============================================================================

EMAIL_HTML = """\
<html>
<body style="font-family: Helvetica, Arial, sans-serif; font-size: 11pt; color: #222; line-height: 1.55;">
<div style="text-align:center; margin-bottom: 18px;">
  <img src="cid:logo" alt="Monique Festival" style="width: 90px; height: auto;">
</div>

<p>Bonjour {prenom},</p>

<p>Ceci est un test d'envoi de reçu fiscal pour valider le modèle que
<strong>Monique Festival</strong> utilisera pour les 14 donateurs du
crowdfunding du 20-23 avril 2026.</p>

<p>Vous trouverez en pièce jointe un <strong>exemple de reçu fiscal</strong> au
format PDF (Cerfa n° 11580*04) pour un don fictif de 50 €.</p>

<p><strong>Points à vérifier dans le document</strong> :</p>
<ul>
  <li>Raison sociale : <strong>MONIQUE FESTIVAL</strong> <em>(anciennement JD Production)</em></li>
  <li>Objet social complet (article 2 des statuts révisés du 19 avril 2026)</li>
  <li>Numéro d'ordre du reçu : <code>{numero_recu}</code></li>
  <li>Article 200 CGI (particuliers) avec réduction de 66 %</li>
  <li>Logo Monique Festival en en-tête du document</li>
  <li>Signature de la Présidente en bas du document</li>
  <li>Mentions obligatoires Cerfa présentes</li>
</ul>

<p><em>Ce reçu ne doit pas être utilisé pour une déclaration fiscale
(mention « RECU TEST » dans le document).</em></p>

<p>Si le modèle vous convient, il sera dupliqué pour les 14 donateurs réels
une fois le récépissé préfectoral reçu (début mai 2026) et les paramètres
Hello Asso mis à jour.</p>

<p>Je reste à votre disposition pour toute remarque ou ajustement avant la
diffusion réelle.</p>

<p>Très cordialement,</p>

<p>
<strong>Judith Laithier</strong><br>
Présidente de l'association Monique Festival<br>
<a href="mailto:info@monique-festival.fr">info@monique-festival.fr</a>
</p>

<p style="font-size: 9pt; color: #888; margin-top: 24px; border-top: 1px dashed #bbb; padding-top: 8px;">
Monique Festival — Festival émergent et pluridisciplinaire à Besançon<br>
28, 29 et 30 août 2026 — La Grange Huguenet<br>
Instagram : <a href="https://instagram.com/monique.festival">@monique.festival</a>
</p>
</body>
</html>
"""


def build_email(destinataire, objet, html_body, pdf_path: Path, logo_path: Path):
    message = MIMEMultipart('related')
    message['To'] = destinataire
    message['Subject'] = objet

    alt = MIMEMultipart('alternative')
    message.attach(alt)
    alt.attach(MIMEText(html_body, 'html', 'utf-8'))

    if logo_path.exists():
        with open(logo_path, 'rb') as f:
            img = MIMEImage(f.read(), _subtype='png')
            img.add_header('Content-ID', '<logo>')
            img.add_header('Content-Disposition', 'inline', filename=logo_path.name)
            message.attach(img)

    if pdf_path.exists():
        with open(pdf_path, 'rb') as f:
            pdf_attach = MIMEBase('application', 'pdf')
            pdf_attach.set_payload(f.read())
        encoders.encode_base64(pdf_attach)
        pdf_attach.add_header(
            'Content-Disposition',
            f'attachment; filename="{pdf_path.name}"'
        )
        message.attach(pdf_attach)
    else:
        print(f"ATTENTION : PDF non trouvé : {pdf_path}")

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw}


def main():
    # 1. Générer PDF
    print("Étape 1/3 — Génération du reçu fiscal PDF (ReportLab)...")
    slug = f"{DON['donateur_prenom']}_{DON['donateur_nom']}".replace(" ", "_").replace("-", "_")
    pdf_path = OUTPUT_DIR / f"TEST_RF_{slug}.pdf"
    build_pdf(DON, pdf_path)
    print(f"  PDF généré : {pdf_path.relative_to(REPO_ROOT)} ({pdf_path.stat().st_size // 1024} KB)")

    # 2. Credentials Gmail
    print("\nÉtape 2/3 — Connexion Gmail API...")
    creds = load_creds()
    service = build('gmail', 'v1', credentials=creds)

    email_body = EMAIL_HTML.format(
        prenom=DON["donateur_prenom"],
        numero_recu=DON["numero_recu"],
    )

    # 3. Envoi
    objet = "[TEST v2] Monique Festival — Exemple de reçu fiscal PDF"
    print(f"\nÉtape 3/3 — Envoi du mail :")
    print(f"  À       : {DON['email']}")
    print(f"  Objet   : {objet}")
    print(f"  PJ      : {pdf_path.name} ({pdf_path.stat().st_size // 1024} KB)")
    print(f"  Logo    : inline (mail) + en-tête (PDF)")
    print(f"  Signature : image intégrée dans le PDF")

    message = build_email(
        destinataire=DON["email"],
        objet=objet,
        html_body=email_body,
        pdf_path=pdf_path,
        logo_path=LOGO_PATH,
    )

    try:
        sent = service.users().messages().send(userId='me', body=message).execute()
        print(f"\nOK — Message ID Gmail : {sent['id']}")
        print(f"Verifie ta boite : {DON['email']}")
    except Exception as e:
        print(f"\nERREUR d'envoi : {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
