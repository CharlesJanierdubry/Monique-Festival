# DocuSign — Envoi automatisé des documents de gouvernance

Scripts Python pour **envoyer, suivre et télécharger** les documents de gouvernance via l'API DocuSign eSignature.

---

## 🎯 Ce que ça permet

- 📤 **Envoyer** une enveloppe DocuSign avec un ou plusieurs signataires à partir des .docx du dossier [`gouvernance/docx/`](../docx/)
- 👀 **Suivre** l'état des enveloppes en cours (envoyé / lu / signé / finalisé)
- 📥 **Télécharger** les documents signés et le certificat d'audit une fois l'enveloppe complétée
- 🔁 **Envois récurrents** : chaque enveloppe est définie dans `envelopes.json`, réutilisable pour les futures conventions de mécénat, contrats artistes, bulletins d'adhésion…

---

## 📦 Prérequis

### 1. Compte DocuSign
- Compte **DocuSign eSignature** (essai gratuit 30 jours ou abonnement personnel / associatif)
- Ou un **compte développeur** gratuit : https://developers.docusign.com

### 2. Credentials API (JWT Grant)
À récupérer sur l'**Admin DocuSign** → **Settings** → **Apps and Keys** :
- **Integration Key** (Client ID)
- **User ID** (API Username)
- **Account ID**
- **RSA Private Key** (à générer et télécharger)

Pour créer une intégration : https://developers.docusign.com/platform/auth/jwt/

### 3. Autorisation JWT (première fois seulement)
Ouvrir dans le navigateur une fois pour autoriser l'intégration :
```
https://account-d.docusign.com/oauth/auth?response_type=code
  &scope=signature%20impersonation
  &client_id={INTEGRATION_KEY}
  &redirect_uri=https://www.docusign.com
```
(remplacer `account-d` par `account` pour un compte de production)

---

## 🛠 Installation

```bash
cd gouvernance/docusign
pip install -r requirements.txt
cp .env.example .env
# Éditer .env avec les credentials DocuSign
# Placer la clé privée RSA dans private.key
```

Le fichier `.env` et `private.key` sont **automatiquement ignorés par Git** (voir [../../.gitignore](../../.gitignore)).

---

## 🚀 Utilisation

### Envoyer une enveloppe
```bash
python send_envelope.py --envelope age_19avril
```

### Lister les enveloppes en cours
```bash
python send_envelope.py --list
```

### Vérifier le statut d'une enveloppe
```bash
python send_envelope.py --status <ENVELOPE_ID>
```

### Télécharger les documents signés
```bash
python send_envelope.py --download <ENVELOPE_ID>
```
Les documents signés sont enregistrés dans `docx_signes/`.

---

## 📋 Configuration des enveloppes

Les enveloppes sont définies dans [`envelopes.json`](envelopes.json). Chaque enveloppe précise :
- `name` : libellé interne
- `subject` : objet de l'email DocuSign au signataire
- `message` : message d'accompagnement
- `documents` : liste des fichiers .docx à faire signer (chemin relatif depuis `gouvernance/`)
- `signers` : liste des signataires avec nom, email, ordre et zones de signature

**4 enveloppes pré-configurées** :
1. `age_19avril` — PV AGE + Statuts révisés (8 signataires)
2. `bureau_post_age` — PV Bureau + Règlement intérieur (Présidente + Trésorier)
3. `mecenat_lip` — Convention de mécénat Lip (modèle, adaptable)
4. `adhesion_nouveau_membre` — Bulletin d'adhésion (modèle)

---

## 🔐 Sécurité

- **Ne jamais commiter** `.env`, `private.key` ou tout fichier contenant des credentials
- Le `.gitignore` à la racine du projet les exclut automatiquement
- Les clés privées RSA doivent être stockées uniquement en local
- En cas de fuite, régénérer immédiatement les clés dans l'Admin DocuSign

---

## 💡 Alternative si tu veux juste un envoi ponctuel

Pour une seule enveloppe (l'AGE par exemple), tu peux aussi utiliser directement l'interface web DocuSign :
1. Se connecter sur app.docusign.com
2. **New** → **Send an Envelope**
3. Upload les .docx depuis `gouvernance/docx/`
4. Ajouter les signataires + positionner les zones de signature à la souris
5. Envoyer

Cela prend ~10 min par enveloppe. Le script devient rentable dès la 3ᵉ enveloppe ou pour les envois récurrents.
