"""
Crée un BROUILLON Gmail à Luther + Lorraine pour les corrections
de la Charte éthique trouvée sur le Drive (version 25/04/2026).

Pas d'envoi — vérification dans Gmail Brouillons.
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
SUBJECT = "Charte éthique — corrections orthographiques + suggestion ajout transport"

CORPS = """Bonjour Luther et Lorraine,

J'ai relu attentivement la Charte éthique sur le Drive (version 25/04/2026, dans Communication & Marketing/Charte éthique.pdf, et copies dans 0_Gouvernance/Charte_ethique/ et 7 contenus infos pratiques/2 Notre éthique/).

J'ai relevé quelques corrections à apporter, dont 2 fautes critiques. Le contenu de fond est très bien — il s'agit uniquement de coquilles et d'incohérences mineures à reprendre dans Canva.

----------------------------------------------------------

CORRECTIONS À APPORTER

== 🔴 Fautes critiques (à corriger absolument) ==

1. PAGE 1 — TITRE D'UNE SECTION
   Affiché : "INCLUSION, RESPECT ET LUTTE CONTRE LES DISCTIMINATIONS"
   Correction : "INCLUSION, RESPECT ET LUTTE CONTRE LES DISCRIMINATIONS"
   (faute à T inversé — particulièrement visible car dans un titre)

2. PAGE 3 — PHRASE DE LA RUBRIQUE "ESPACE RESSOURCES ET INCLUSION"
   Affiché : "Des stands d'association seront présent sur le festival."
   Correction : "Des stands d'associations seront présents sur le festival."
   (deux fautes : "associations" au pluriel + "présents" au pluriel — accord avec "stands")

== 🟠 Espace manquant (page 1) ==

Affiché : "...sûr, inclusif, responsable et bienveillant pour toutes et tous.Toute personne participant..."
Correction : "...toutes et tous. Toute personne participant..."
(manque un espace après le point)

== 🟡 Cohérence écriture inclusive ==

Vous utilisez parfois "chacun·e", "attentif·ve", "participant·e", parfois non. À harmoniser :

- PAGE 2 (engagement bénévoles) : "Être attentifs, disponibles" → "Être attentif·ves, disponibles"
- PAGE 4 (engagements écologiques) : "Chaque participant est responsable" → "Chaque participant·e est responsable"

== 🟡 Tournure à reformuler (page 2, engagement artistes) ==

Affiché : "Ne tenir aucun propos ou comportement discriminatoire ou violent"
Suggestion : "Ne tenir aucun propos discriminatoire ou violent et ne pas adopter de comportement de cette nature"
(on tient un propos, on adopte un comportement — séparer rend la phrase plus claire)

== 🟡 Détail mineur (page 2) ==

"mal intentionné" → "malintentionné" (l'Académie française recommande la version en un seul mot)

----------------------------------------------------------

SUGGESTION D'AJOUT — TRANSPORT (cohérence avec la règle de défraiement)

La rubrique "Engagements écologiques" liste 5 engagements (déchets, mégots, plastique, nature, ne rien laisser) mais ne mentionne pas le transport. Or nous venons de finaliser une règle de défraiement qui exclut explicitement la voiture personnelle (cohérent avec la mobilité douce).

Pour renforcer la cohérence Charte ↔ Règle de défraiement, je suggère d'ajouter dans "Engagements écologiques" :

"Privilégier les modes de transport doux pour se rendre au festival (train, autocar, vélo, covoiturage)."

Cela nous permettra aussi d'argumenter solidement vis-à-vis des artistes étrangers qui se demanderaient pourquoi on ne défraie pas la voiture.

----------------------------------------------------------

PROCHAINE ÉTAPE

Idéalement, ces corrections devraient être intégrées dans Canva avant la prochaine vague de diffusion (signature artistes, mécènes, partenaires).

Si vous voulez, je peux aussi vous proposer une version corrigée du texte sous forme de Markdown que vous n'auriez plus qu'à recopier dans Canva.

Bien à vous,

Charles
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
    print(f"Brouillon Gmail créé.")
    print(f"  Draft ID  : {draft['id']}")
    print(f"  De        : {FROM}")
    print(f"  À         : {TO}")
    print(f"  Objet     : {SUBJECT}")


if __name__ == "__main__":
    main()
