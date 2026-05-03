"""
Cree 2 brouillons Gmail vers AIAC Courtage + MAIF (pour devis assurance RC pro)
+ produit un info pack pour le formulaire en ligne Smacl.

Pieces jointes :
 - Recepisse_modification_24-04-2026.pdf
 - Statuts_Monique_Festival_19-04-2026_Signe.pdf
 - Budget_festival_recettes.md (converti en PDF si possible, sinon .md)

Usage : python scripts/draft_assurance_devis_3.py
"""
import sys, os, json, base64
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN_FILE = os.path.expanduser('~/.config/mcp-gdrive/token.json')
REPO = Path(__file__).resolve().parent.parent

# Pieces jointes
RECEPISSE = REPO / "officiel/prefecture/Recepisse_modification_24-04-2026.pdf"
STATUTS = REPO / "gouvernance/docx/Statuts_Monique_Festival_19-04-2026_Signé.pdf"
PV_AGE = REPO / "gouvernance/docx/PV_AGE_adoption_statuts 19-04-2026 - Signé.pdf"

# Pour les destinataires
DESTINATAIRES = [
    {
        "label": "AIAC Courtage",
        "email": "contact@aiac.fr",
        "tel": "01 44 53 28 53",
        "specifite": "Spécialiste festivals et événements culturels",
    },
    {
        "label": "MAIF Associations",
        "email": "contact@maif.fr",
        "tel": "09 78 97 98 99",
        "specifite": "Bon rapport qualité-prix pour petites/moyennes asso",
    },
]


CORPS_TEMPLATE = """Bonjour,

Je sollicite un devis d'assurance Responsabilité Civile professionnelle (et garanties associées) pour notre association culturelle qui organise sa première édition de festival pluridisciplinaire à Besançon.

Nous comparons plusieurs courtiers/assureurs spécialisés et serions intéressés par votre proposition.

============================================
IDENTITÉ DE L'ASSOCIATION
============================================

- Dénomination : MONIQUE FESTIVAL (anciennement JD Production, mod. statuts AGE 19/04/2026, récépissé Préfecture du Doubs 24/04/2026)
- Forme juridique : Association Loi 1901 — d'intérêt général à caractère culturel
- N° RNA : W251011013
- SIREN : 991 055 450
- SIRET : 991 055 450 00019
- Code APE : 93.29Z (en cours de modification vers 90.01Z « Arts du spectacle vivant »)
- Siège : 54 chemin de Valentin, 25000 Besançon
- Présidente (représentante légale) : Judith LAITHIER
- Trésorier : Wadih CORMIER

============================================
LE FESTIVAL — ÉDITION 2026 (1ʳᵉ ÉDITION)
============================================

- Intitulé : Monique Festival — Festival pluridisciplinaire émergent
- Dates : 28, 29 et 30 août 2026 (3 jours consécutifs)
- Lieu : La Grange Huguenet, Besançon — bâtisse classée aux monuments historiques, dans un parc arboré de 4 hectares
- Régime de location : convention avec le propriétaire (le bail est en cours de mise à jour avec la nouvelle dénomination)
- Programmation : musique classique, musiques actuelles, théâtre — environ 30 artistes
- Nombre de scènes : 2 (scène amplifiée + scène acoustique)
- Action culturelle : 4 ateliers (chant lyrique, écriture, musique & théâtre, DJ) avec restitution publique
- Jauge prévisionnelle : 1 200 festivaliers cumulés sur les 3 jours (~400/jour)
- Bénévoles attendus : ~70

============================================
ACTIVITÉS ÉCONOMIQUES ANNEXES
============================================

- Billetterie en ligne (Hello Asso) et sur place (iPhone Tap to Pay / SumUp)
- Restauration sur site (producteurs locaux, environ 1 000 sandwichs/3j)
- Buvette avec licence III temporaire (bière, vin, soft) — gobelets réutilisables consignés
- Marché d'artisans / animations (photo-booth, atelier maquillage, échecs, etc.)

============================================
BUDGET PRÉVISIONNEL DE L'ÉDITION 2026
============================================

- Recettes prévisionnelles : ~58 500 € (billetterie 22 650 € + buvette 20 720 € + restauration 11 722 € + ateliers 3 390 €)
- Dépenses prévisionnelles : ~49 800 €
- Marge prévisionnelle : ~+8 600 €
- Cachets artistes (grille uniforme GUSO) : ~15 700 € (tous artistes via GUSO)

============================================
GARANTIES SOUHAITÉES
============================================

- Responsabilité Civile organisateur d'événement
- Responsabilité Civile exploitation
- Dommages aux biens loués (notamment La Grange Huguenet — site classé)
- Vol et vandalisme du matériel sur site
- Couverture des bénévoles (~70 personnes)
- Protection juridique (souhaitable)
- Annulation événement : optionnelle, à chiffrer séparément si possible

Plafond de garantie souhaité : 1 M€ par sinistre minimum.
Franchise raisonnable acceptée.

============================================
PIÈCES JOINTES
============================================

1. Récépissé préfectoral de modification (24/04/2026)
2. Statuts révisés signés (19/04/2026)
3. Procès-verbal de l'AGE du 19/04/2026 (signé)

Je peux compléter avec :
- Plan de site (en cours de finalisation)
- Charte éthique du festival
- Programmation détaillée des 3 jours
- Budget prévisionnel intégral

============================================
DEMANDE
============================================

Pourriez-vous m'adresser un devis détaillé sous 5-10 jours ouvrés ?

Notre échéance interne : nous souhaitons souscrire avant le 20 mai 2026 pour pouvoir produire l'attestation aux mécènes, à la commission de sécurité, à La Grange Huguenet et à notre prestataire technique Régis Régis.

Je suis disponible pour un échange téléphonique ou en visio à votre convenance.

Bien cordialement,

Charles JANIER-DUBRY
Pôle Administratif — Monique Festival
charles@janier-dubry.fr
07 87 43 87 85

Pour le Bureau (Présidente Judith LAITHIER, Trésorier Wadih CORMIER)
"""


def load_creds():
    with open(TOKEN_FILE) as f:
        d = json.load(f)
    if 'https://www.googleapis.com/auth/gmail.compose' not in d.get('scopes', []):
        sys.exit("ERREUR : scope gmail.compose manquant. Lance python scripts/gmail_auth.py")
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid: creds.refresh(Request())
    return creds


def ascii_filename(name):
    import unicodedata, re
    nfkd = unicodedata.normalize('NFKD', name)
    no_accents = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r'\s+', '_', no_accents)


def build_message(dest, objet, corps, pjs):
    msg = MIMEMultipart()
    msg['From'] = 'Monique Festival <info@monique-festival.fr>'
    msg['To'] = dest
    msg['Subject'] = objet
    msg.attach(MIMEText(corps, 'plain', 'utf-8'))
    for p in pjs:
        with open(p, 'rb') as f:
            attach = MIMEBase('application', 'pdf')
            attach.set_payload(f.read())
        encoders.encode_base64(attach)
        safe = ascii_filename(p.name)
        attach.add_header('Content-Disposition', f'attachment; filename="{safe}"')
        msg.attach(attach)
    return base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')


def main():
    pjs = [RECEPISSE, STATUTS, PV_AGE]
    print("Pieces jointes :")
    for p in pjs:
        if not p.exists():
            sys.exit(f"ERREUR : {p} absente")
        print(f"  OK  {p.name} ({p.stat().st_size // 1024} KB)")

    gmail = build('gmail', 'v1', credentials=load_creds())

    print("\nCreation des brouillons Gmail :")
    for d in DESTINATAIRES:
        objet = f"Demande de devis assurance RC Pro — Festival Monique 2026 (1ʳᵉ édition, 28-30 août 2026, Besançon)"
        raw = build_message(d['email'], objet, CORPS_TEMPLATE, pjs)
        draft = gmail.users().drafts().create(userId='me', body={'message': {'raw': raw}}).execute()
        print(f"  OK  {d['label']} ({d['email']}) - Draft ID : {draft['id']}")

    # Sauvegarde de l'info pack pour Smacl (formulaire en ligne)
    info_pack = REPO / "officiel" / "logs" / "Info_pack_assurance_devis.txt"
    info_pack.parent.mkdir(parents=True, exist_ok=True)
    info_pack.write_text(
        "INFO PACK - DEMANDE DEVIS ASSURANCE RC PRO MONIQUE FESTIVAL\n"
        "=" * 70 + "\n\n"
        "Pour le formulaire SMACL en ligne :\n"
        "  https://services.smacl.fr/devis-assurance-association/\n"
        "ou par tel : 05 49 32 34 96 (asso sans employe) / 05 49 34 29 30 (asso avec employe)\n\n"
        "=" * 70 + "\n"
        "TEXTE DESCRIPTIF A COLLER DANS LE CHAMP \"DESCRIPTION\"\n"
        "=" * 70 + "\n\n"
        + CORPS_TEMPLATE,
        encoding='utf-8',
    )
    print(f"\nInfo pack pour SMACL (formulaire en ligne) : {info_pack.relative_to(REPO)}")

    print("\n=== Termine ===")
    print(f"\n2 brouillons Gmail crees :")
    for d in DESTINATAIRES:
        print(f"  - {d['label']} ({d['email']}) - tel {d['tel']} - {d['specifite']}")
    print(f"\nFormulaire SMACL en ligne : https://services.smacl.fr/devis-assurance-association/")
    print(f"  Telephone : 05 49 32 34 96")
    print(f"  Texte pret a coller : {info_pack.relative_to(REPO)}")


if __name__ == "__main__":
    main()
