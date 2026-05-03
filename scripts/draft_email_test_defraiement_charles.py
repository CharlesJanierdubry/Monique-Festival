"""
Crée un BROUILLON Gmail HTML soigné, envoyé en test à charles@janier-dubry.fr,
pour valider le rendu avant envoi aux 6 artistes étrangers.

L'email :
- Logo Monique Festival inline (en-tête)
- Texte EN soigné
- Tableau récap des conditions
- Lien vers le PDF Drive
- Procédure d'acceptation claire (email reply)

Pas d'envoi direct — vérification dans Gmail Brouillons.
"""
import sys, os, json, base64
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO = Path(__file__).resolve().parent.parent
LOGO = REPO / "Logo_Monique_Festival.png"

FROM = 'Monique Festival <info@monique-festival.fr>'
TO = 'charles@janier-dubry.fr'
SUBJECT = "[TEST] Monique Festival 2026 — Travel reimbursement (please prepare)"

PDF_LINK = "https://drive.google.com/file/d/1_9pvDMC1e-kSeZEcqG45ogVzMpDuaftr/view"

HTML_BODY = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          color: #2c2c2c; line-height: 1.5; max-width: 640px; margin: 0 auto; }}
  h2 {{ color: #1F4E78; border-bottom: 2px solid #E07A5F; padding-bottom: 6px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 16px 0; font-size: 14px; }}
  th, td {{ border: 1px solid #d0d0d0; padding: 8px 12px; text-align: left; vertical-align: top; }}
  th {{ background-color: #1F4E78; color: white; }}
  .accept-box {{ background-color: #FFF2CC; border-left: 4px solid #E07A5F;
                 padding: 12px 18px; margin: 18px 0; border-radius: 4px; }}
  .accept-box strong {{ color: #1F4E78; }}
  .button {{ display: inline-block; background-color: #1F4E78; color: white !important;
             padding: 10px 18px; text-decoration: none; border-radius: 4px;
             font-weight: bold; margin: 10px 0; }}
  ul li {{ margin-bottom: 6px; }}
  .footer {{ font-size: 12px; color: #777; border-top: 1px solid #ddd;
             padding-top: 12px; margin-top: 24px; }}
</style>
</head>
<body>

<p style="text-align: center; margin: 20px 0;">
  <img src="cid:logo_monique" alt="Monique Festival" width="120" style="display: inline-block;">
</p>

<p>Hi <strong>[Artist first name]</strong>,</p>

<p>We're really happy to welcome you at <strong>Monique Festival</strong> on <strong>28-30 August 2026</strong> in Besançon, France.</p>

<p>Since you're traveling from abroad, here is everything you need to know about our travel reimbursement, fully aligned with our environmental charter.</p>

<h2>📋 Reimbursement at a glance</h2>

<table>
  <tr><th>Item</th><th>Detail</th></tr>
  <tr><td><strong>Maximum cap</strong></td><td>€200 per person, refunded on actual cost (proof needed)</td></tr>
  <tr><td><strong>Eligible transport</strong></td><td>Train, long-distance bus, plane (only if no rail option)</td></tr>
  <tr><td><strong>NOT eligible</strong></td><td>Personal car (eco charter)</td></tr>
  <tr><td><strong>Meals + accommodation</strong></td><td>Provided in kind on-site at La Grange Huguenet</td></tr>
  <tr><td><strong>Booking deadline</strong></td><td>Before <strong>15 June 2026</strong> (prices spike after)</td></tr>
  <tr><td><strong>Refund delay</strong></td><td>Within 7 days of receiving your documents</td></tr>
</table>

<h2>📄 Full agreement</h2>

<p>Please read the complete Travel Reimbursement Agreement (1 page, EN):</p>

<p style="text-align: center;">
  <a href="{pdf_link}" class="button">Open the PDF agreement</a>
</p>

<h2>✅ How to accept &amp; what to send us</h2>

<div class="accept-box">
<p><strong>Reply to this email</strong> at <a href="mailto:info@monique-festival.fr">info@monique-festival.fr</a> with:</p>
<ol>
  <li>The exact sentence:<br>
      <em>« I, [your full name], have read and accept the Travel Reimbursement Agreement of Monique Festival 2026. »</em></li>
  <li>A screenshot or PDF of your <strong>RIB / IBAN + BIC</strong></li>
  <li>(As soon as you book) a screenshot or PDF of your <strong>travel ticket</strong></li>
</ol>
<p>That's it — no need for paper signature or DocuSign (available on request if you prefer).</p>
</div>

<p>Looking forward to having you with us in Besançon!</p>

<p>Best regards,<br>
<strong>Charles JANIER-DUBRY</strong><br>
Administrative Lead — Monique Festival<br>
<a href="mailto:info@monique-festival.fr">info@monique-festival.fr</a> · +33 7 87 43 87 85</p>

<div class="footer">
  Monique Festival is a French non-profit association under the 1901 Act.<br>
  RNA W251011013 · SIREN 991 055 450 · 54 chemin de Valentin, 25000 Besançon, France
</div>

</body>
</html>
""".format(pdf_link=PDF_LINK)


PLAIN_FALLBACK = """Hi [Artist first name],

We're really happy to welcome you at Monique Festival on 28-30 August 2026 in Besançon, France.

== Reimbursement at a glance ==
- Maximum cap: €200 per person, on actual cost (proof needed)
- Eligible transport: train, long-distance bus, plane (only if no rail option)
- Not eligible: personal car (eco charter)
- Meals + accommodation: provided in kind on-site at La Grange Huguenet
- Booking deadline: before 15 June 2026
- Refund delay: within 7 days

== Full agreement ==
{pdf_link}

== How to accept & what to send us ==
Reply to this email (info@monique-festival.fr) with:
1. The exact sentence: "I, [your full name], have read and accept the Travel Reimbursement Agreement of Monique Festival 2026."
2. Screenshot or PDF of your RIB/IBAN + BIC
3. (As soon as you book) screenshot or PDF of your travel ticket

Best regards,
Charles JANIER-DUBRY
Administrative Lead — Monique Festival
info@monique-festival.fr · +33 7 87 43 87 85
""".format(pdf_link=PDF_LINK)


def main():
    if not LOGO.exists():
        sys.exit(f"Logo introuvable : {LOGO}")
    with open(TOKEN) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid:
        creds.refresh(Request())
    gmail = build('gmail', 'v1', credentials=creds)

    # Structure : multipart/related (HTML + image inline)
    #   └── multipart/alternative
    #         ├── text/plain (fallback)
    #         └── text/html
    #   └── image (inline, cid:logo_monique)
    msg = MIMEMultipart('related')
    msg['From'] = FROM
    msg['To'] = TO
    msg['Subject'] = SUBJECT

    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(PLAIN_FALLBACK, 'plain', 'utf-8'))
    alt.attach(MIMEText(HTML_BODY, 'html', 'utf-8'))
    msg.attach(alt)

    with open(LOGO, 'rb') as f:
        img = MIMEImage(f.read(), 'png')
    img.add_header('Content-ID', '<logo_monique>')
    img.add_header('Content-Disposition', 'inline', filename='Logo_Monique_Festival.png')
    msg.attach(img)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
    draft = gmail.users().drafts().create(
        userId='me', body={'message': {'raw': raw}}).execute()
    print(f"Brouillon HTML créé.")
    print(f"  Draft ID  : {draft['id']}")
    print(f"  De        : {FROM}")
    print(f"  À         : {TO}")
    print(f"  Objet     : {SUBJECT}")
    print(f"\nVérifie dans Gmail (festivalljd@) > Brouillons.")
    print(f"Une fois validé, je pourrai dupliquer pour les 6 artistes (avec personnalisation prénom).")


if __name__ == "__main__":
    main()
