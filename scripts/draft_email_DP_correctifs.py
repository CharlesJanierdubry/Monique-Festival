"""
Crée un BROUILLON Gmail à Luther + Lorraine (co-chefs Communication)
pour demander le lien Canva du DP + déposer dans le Drive
+ liste des correctifs à apporter.

Pas d'envoi — vérification manuelle dans Gmail Brouillons.
"""
import sys, os, json, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TOKEN = os.path.expanduser('~/.config/mcp-gdrive/token.json')

FROM = 'Monique Festival <info@monique-festival.fr>'
TO = 'lututersan12@gmail.com, lorrainejanierdubry@gmail.com'
SUBJECT = "Dossier de presse — récupération du lien Canva + correctifs à apporter"

CORPS = """Bonjour Luther et Lorraine,

J'espère que vous allez bien.

Je viens de relire le dossier de presse "Mi-avril 2026" qui est sur notre Drive (Communication & Marketing > Dossier de presse > Mi-avril dossier de presse.pdf) et j'ai relevé plusieurs corrections à faire avant la prochaine vague de diffusion presse / partenaires / mécènes.

Or je n'arrive pas à retrouver le fichier source Canva sur le Drive — seul le PDF y est archivé. Pourriez-vous :

1. Partager le lien Canva du DP avec moi (charles@janier-dubry.fr ou info@monique-festival.fr) — accès édition de préférence
2. Déposer ce lien dans le dossier Drive "Communication & Marketing > Dossier de presse/" sous forme d'un Google Doc nommé par exemple "Lien_Canva_DP_Mi-avril.gdoc" — ainsi tout futur intervenant pourra retrouver la source directement à côté du PDF

----------------------------------------------------------

SYNTHÈSE DES CORRECTIFS À APPORTER SUR LE DOSSIER DE PRESSE

== Fautes d'orthographe / grammaire (6) ==

- "Festival émergeant" → "Festival émergent" (3 occurrences : titre p.1, "Pour qui ?" p.3, "Notre Vision Artistique" p.5). « Émergent » est l'adjectif ; « émergeant » est le participe présent du verbe.
- "Nous sommes sensibles de créer" → "sensibles à créer" (rubrique Discrimination)
- "Bénévoles recrutés par leur compétences" → "pour leurs compétences" (rubrique Discrimination — double erreur)
- "où se rencontrent générations, disciplines et sensibilités" → ajouter les articles : "où se rencontrent les générations, les disciplines..."
- "chacun·ne" → "chacun·e" (forme inclusive standard, rubrique Écologie)
- "toutes formes de violences" → "toutes les formes de violence" (article + singulier générique)

== Erreurs factuelles à corriger absolument (3) ==

- "Voiture" classée mobilité douce (rubrique Écologie) — c'est faux et risqué pour la presse écolo. Reformuler en : "défrayés uniquement pour des trajets en transports en commun (bus, train) ou en covoiturage"
- "Fratrie de cinq artistes" mentionnée 2 fois (Pour qui + Histoire), mais l'équipe fondatrice ne liste que 4 artistes (Lorraine, Léopold, Lewis, Luther). À trancher : soit "fratrie de quatre", soit ajouter le 5e. ATTENTION : Léontine ne doit pas figurer dans des documents publics (mineure).
- Email de contact incohérent : page 1 = info@monique-festival.fr (correct), page Contact = info@festival-monique.fr (mauvaise URL inversée). À uniformiser sur info@monique-festival.fr.

== Tournures à reformuler (qualité éditoriale) ==

- Phrase fragmentée rubrique "Notre Vision Artistique" — les phrases commençant par "En proposant..." et "Afin de..." sont sans verbe principal, à fondre dans une seule phrase
- "L'accessibilité pour tous et toutes est ce que nous voulons afin de..." — formulation lourde, à reformuler
- "vision la plus verte possible" (Écologie) — une vision n'est pas verte. Préférer "démarche la plus écologique possible"
- "pour la rareté de ses granges urbaines" (Le Lieu) — la Grange Huguenet est une seule grange. Reformuler en "la rareté de cette grange en milieu urbain"
- "notre attention en sera toute particulière quant au soin de chacun·ne à prendre soin de notre environnement" — phrase bancale, à reformuler

== Détails mineurs ==

- "4.7 / Historical landmark" (Le Lieu) — anglais collé depuis Google Maps, à traduire : "Note 4,7/5 — Monument historique"
- Capitalisation incohérente : "Monique festival" (f minuscule) vs "Monique Festival" (F majuscule) — à uniformiser

----------------------------------------------------------

Pas d'urgence absolue, mais idéalement avant la prochaine vague de diffusion (ouverture billetterie mi-juin, démarchage mécènes juin-juillet).

Si vous préférez, je peux faire les corrections moi-même une fois le lien Canva en accès édition — dites-moi.

Bien à vous,

Charles JANIER-DUBRY
Pôle Administratif — Monique Festival
charles@janier-dubry.fr
"""


def main():
    with open(TOKEN) as f:
        d = json.load(f)
    creds = Credentials(token=d['token'], refresh_token=d.get('refresh_token'),
        token_uri=d['token_uri'], client_id=d['client_id'],
        client_secret=d['client_secret'], scopes=d['scopes'])
    if not creds.valid:
        creds.refresh(Request())
    gmail = build('gmail', 'v1', credentials=creds)

    msg = MIMEMultipart()
    msg['From'] = FROM
    msg['To'] = TO
    msg['Subject'] = SUBJECT
    msg.attach(MIMEText(CORPS, 'plain', 'utf-8'))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')

    draft = gmail.users().drafts().create(
        userId='me', body={'message': {'raw': raw}}).execute()
    print(f"Brouillon créé.")
    print(f"  Draft ID  : {draft['id']}")
    print(f"  De        : {FROM}")
    print(f"  À         : {TO}")
    print(f"  Objet     : {SUBJECT}")
    print(f"\nVérifie dans Gmail (festivalljd@) > Brouillons.")
    print(f"Avant envoi, contrôle bien que le champ 'De' affiche info@monique-festival.fr.")


if __name__ == "__main__":
    main()
