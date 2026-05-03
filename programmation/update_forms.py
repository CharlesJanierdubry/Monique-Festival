"""
Met à jour les formulaires Google Forms : ajout champs manquants.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import json, os

with open(os.path.expanduser('~/.config/mcp-gdrive/token.json')) as f:
    td = json.load(f)
creds = Credentials(
    token=td['token'], refresh_token=td['refresh_token'],
    token_uri=td['token_uri'], client_id=td['client_id'],
    client_secret=td['client_secret'], scopes=td['scopes']
)
forms_service = build('forms', 'v1', credentials=creds)

# === FORMULAIRE INDIVIDUEL ===
FORM_INDIV = '1kQQApVrh5xbqQND6JRmYbPEfAD-_viBk0o9JSZ2qPm4'

form = forms_service.forms().get(formId=FORM_INDIV).execute()
items = form.get('items', [])

secu_item_id = items[3]['itemId']
secu_q_id = items[3]['questionItem']['question']['questionId']

requests = [
    # Rendre N secu obligatoire
    {
        'updateItem': {
            'item': {
                'itemId': secu_item_id,
                'title': 'Numero de securite sociale',
                'questionItem': {
                    'question': {
                        'questionId': secu_q_id,
                        'required': True,
                        'textQuestion': {}
                    }
                }
            },
            'location': {'index': 3},
            'updateMask': 'questionItem.question.required'
        }
    },
    # Nom de scene
    {
        'createItem': {
            'item': {
                'title': 'Nom de scene / nom du projet',
                'questionItem': {
                    'question': {
                        'required': True,
                        'textQuestion': {}
                    }
                }
            },
            'location': {'index': 2}
        }
    },
]

result = forms_service.forms().batchUpdate(formId=FORM_INDIV, body={'requests': requests}).execute()
print(f"INDIVIDUEL phase 1: {len(result.get('replies', []))} modifications OK")

# Phase 2: ajouter charte et RGPD (après les insertions)
form2 = forms_service.forms().get(formId=FORM_INDIV).execute()
nb_items = len(form2.get('items', []))

requests2 = [
    {
        'createItem': {
            'item': {
                'title': 'Charte ethique du Monique Festival',
                'description': "J'ai lu et je m'engage a respecter la charte ethique du Monique Festival 2026.",
                'questionItem': {
                    'question': {
                        'required': True,
                        'choiceQuestion': {
                            'type': 'CHECKBOX',
                            'options': [{'value': "Oui, j'accepte la charte ethique"}]
                        }
                    }
                }
            },
            'location': {'index': nb_items}
        }
    },
    {
        'createItem': {
            'item': {
                'title': 'Consentement RGPD',
                'description': "En cochant cette case, vous etes informe(e) que JD Production collecte et traite vos donnees personnelles dans le cadre exclusif de l'organisation du Monique Festival 2026. Vos donnees seront conservees pendant 5 ans. Conformement au RGPD, vous disposez d'un droit d'acces, de rectification et de suppression en contactant JD Production.",
                'questionItem': {
                    'question': {
                        'required': True,
                        'choiceQuestion': {
                            'type': 'CHECKBOX',
                            'options': [{'value': "J'accepte le traitement de mes donnees personnelles"}]
                        }
                    }
                }
            },
            'location': {'index': nb_items + 1}
        }
    },
]
result2 = forms_service.forms().batchUpdate(formId=FORM_INDIV, body={'requests': requests2}).execute()
print(f"INDIVIDUEL phase 2: {len(result2.get('replies', []))} modifications OK")


# === FORMULAIRE COLLECTIF ===
FORM_COLL = '1FEgWEsvpnnxcBlMLiYKLPYaFjjg7YEY7Ux5QBO0Iwb8'

form_c = forms_service.forms().get(formId=FORM_COLL).execute()
items_c = form_c.get('items', [])

rib_c_item_id = items_c[9]['itemId']
rib_c_q_id = items_c[9]['questionItem']['question']['questionId']

requests_c = [
    # Note: RIB file upload ne peut pas etre modifie via API, a rendre obligatoire manuellement
    # SIRET
    {
        'createItem': {
            'item': {
                'title': 'N SIRET ou RNA de la structure',
                'questionItem': {
                    'question': {
                        'required': True,
                        'textQuestion': {}
                    }
                }
            },
            'location': {'index': 9}
        }
    },
    # Licence
    {
        'createItem': {
            'item': {
                'title': 'N licence entrepreneur de spectacles (ou indiquez moins de 6 representations/an)',
                'questionItem': {
                    'question': {
                        'required': True,
                        'textQuestion': {}
                    }
                }
            },
            'location': {'index': 10}
        }
    },
    # Convention collective
    {
        'createItem': {
            'item': {
                'title': 'Convention collective applicable (si connue)',
                'questionItem': {
                    'question': {
                        'required': False,
                        'textQuestion': {}
                    }
                }
            },
            'location': {'index': 11}
        }
    },
]

result_c = forms_service.forms().batchUpdate(formId=FORM_COLL, body={'requests': requests_c}).execute()
print(f"COLLECTIF phase 1: {len(result_c.get('replies', []))} modifications OK")

# Phase 2: charte + RGPD
form_c2 = forms_service.forms().get(formId=FORM_COLL).execute()
nb_c = len(form_c2.get('items', []))

requests_c2 = [
    {
        'createItem': {
            'item': {
                'title': 'Charte ethique du Monique Festival',
                'description': "Le collectif et chaque artiste du collectif s'engagent a lire et respecter la charte ethique du Monique Festival 2026.",
                'questionItem': {
                    'question': {
                        'required': True,
                        'choiceQuestion': {
                            'type': 'CHECKBOX',
                            'options': [{'value': "Oui, le collectif et chaque artiste acceptent la charte ethique"}]
                        }
                    }
                }
            },
            'location': {'index': nb_c}
        }
    },
    {
        'createItem': {
            'item': {
                'title': 'Consentement RGPD',
                'description': "En cochant cette case, vous etes informe(e) que JD Production collecte et traite les donnees personnelles du collectif et de ses membres dans le cadre exclusif de l'organisation du Monique Festival 2026. Les donnees seront conservees pendant 5 ans. Conformement au RGPD, vous disposez d'un droit d'acces, de rectification et de suppression en contactant JD Production.",
                'questionItem': {
                    'question': {
                        'required': True,
                        'choiceQuestion': {
                            'type': 'CHECKBOX',
                            'options': [{'value': "J'accepte le traitement des donnees personnelles"}]
                        }
                    }
                }
            },
            'location': {'index': nb_c + 1}
        }
    },
]
result_c2 = forms_service.forms().batchUpdate(formId=FORM_COLL, body={'requests': requests_c2}).execute()
print(f"COLLECTIF phase 2: {len(result_c2.get('replies', []))} modifications OK")
