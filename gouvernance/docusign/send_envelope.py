"""
DocuSign — Envoi automatisé des documents de gouvernance du Monique Festival.

Usage :
    python send_envelope.py --envelope <KEY>        Envoie une enveloppe définie dans envelopes.json
    python send_envelope.py --list                  Liste les enveloppes en cours chez DocuSign
    python send_envelope.py --status <ENVELOPE_ID>  Vérifie le statut d'une enveloppe
    python send_envelope.py --download <ENVELOPE_ID>  Télécharge les documents signés

Configuration : voir README.md et .env.example
"""

import argparse
import base64
import json
import os
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

from docusign_esign import (
    ApiClient,
    EnvelopeDefinition,
    EnvelopesApi,
    Document,
    Signer,
    SignHere,
    Tabs,
    Recipients,
)
from docusign_esign.client.api_exception import ApiException
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[2]  # racine du repo
GOUV = ROOT / "gouvernance"
ENV_FILE = Path(__file__).parent / ".env"
ENVELOPES_FILE = Path(__file__).parent / "envelopes.json"
SIGNED_DIR = Path(__file__).parent / "docx_signes"

# === Chargement de l'environnement ===
load_dotenv(ENV_FILE)

DS_ENV = os.getenv("DOCUSIGN_ENV", "demo")
INTEGRATION_KEY = os.getenv("DOCUSIGN_INTEGRATION_KEY")
USER_ID = os.getenv("DOCUSIGN_USER_ID")
ACCOUNT_ID = os.getenv("DOCUSIGN_ACCOUNT_ID")
PRIVATE_KEY_PATH = os.getenv("DOCUSIGN_PRIVATE_KEY_PATH", "./private.key")

AUTH_HOST = "account-d.docusign.com" if DS_ENV == "demo" else "account.docusign.com"
BASE_PATH = "https://demo.docusign.net/restapi" if DS_ENV == "demo" else "https://www.docusign.net/restapi"


def _check_env():
    missing = [k for k, v in {
        "DOCUSIGN_INTEGRATION_KEY": INTEGRATION_KEY,
        "DOCUSIGN_USER_ID": USER_ID,
        "DOCUSIGN_ACCOUNT_ID": ACCOUNT_ID,
    }.items() if not v]
    if missing:
        sys.exit(f"❌ Variables manquantes dans .env : {', '.join(missing)}")
    key_path = (Path(__file__).parent / PRIVATE_KEY_PATH).resolve() if not Path(PRIVATE_KEY_PATH).is_absolute() else Path(PRIVATE_KEY_PATH)
    if not key_path.exists():
        sys.exit(f"❌ Clé privée RSA introuvable : {key_path}")
    return key_path


def _get_api_client() -> ApiClient:
    """Authentification JWT Grant auprès de DocuSign."""
    key_path = _check_env()
    client = ApiClient()
    with open(key_path, "rb") as f:
        private_key = f.read()
    token = client.request_jwt_user_token(
        client_id=INTEGRATION_KEY,
        user_id=USER_ID,
        oauth_host_name=AUTH_HOST,
        private_key_bytes=private_key,
        expires_in=3600,
        scopes=["signature", "impersonation"],
    )
    # Récupérer la base_uri spécifique au compte depuis user_info
    user_info = client.get_user_info(token.access_token)
    account = next((a for a in user_info.accounts if a.account_id == ACCOUNT_ID), None)
    if account is None:
        account = next((a for a in user_info.accounts if a.is_default), user_info.accounts[0])
    base_path = f"{account.base_uri}/restapi"
    client.host = base_path
    client.set_default_header("Authorization", f"Bearer {token.access_token}")
    return client


def _load_envelope_config(key: str) -> dict:
    if not ENVELOPES_FILE.exists():
        sys.exit(f"❌ {ENVELOPES_FILE} introuvable.")
    data = json.loads(ENVELOPES_FILE.read_text(encoding="utf-8"))
    if key not in data:
        sys.exit(f"❌ Enveloppe inconnue : {key}. Disponibles : {', '.join(data.keys())}")
    return data[key]


def _build_document(doc_path: Path, doc_id: int) -> Document:
    if not doc_path.exists():
        sys.exit(f"❌ Document introuvable : {doc_path}")
    with open(doc_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()
    return Document(
        document_base64=content_b64,
        name=doc_path.stem,
        file_extension=doc_path.suffix.lstrip("."),
        document_id=str(doc_id),
    )


def _build_signers(signer_cfgs: list, docs: list) -> list:
    """Construit les signataires avec leurs ancres respectives (per-signer anchors)."""
    signers = []
    for idx, s in enumerate(signer_cfgs, start=1):
        anchors = s.get("anchors") or ["Signature :"]
        sign_here_tabs = [
            SignHere(
                anchor_string=a,
                anchor_x_offset="0",
                anchor_y_offset="0",
                anchor_units="inches",
                anchor_ignore_if_not_present="true",
            )
            for a in anchors
        ]
        signer = Signer(
            email=s["email"],
            name=s["name"],
            recipient_id=str(idx),
            routing_order=str(s.get("order", idx)),
            tabs=Tabs(sign_here_tabs=sign_here_tabs),
        )
        signers.append(signer)
    return signers


def send_envelope(envelope_key: str):
    cfg = _load_envelope_config(envelope_key)
    print(f"📤 Préparation de l'enveloppe : {cfg['name']}")
    client = _get_api_client()

    # Documents
    docs = []
    for i, rel in enumerate(cfg["documents"], start=1):
        path = ROOT / rel if not rel.startswith("gouvernance/") else GOUV / rel.split("gouvernance/", 1)[1]
        # Uniformiser les chemins relatifs au repo
        path = (ROOT / rel).resolve()
        docs.append(_build_document(path, i))
        print(f"   📄 {path.name}")

    signers = _build_signers(cfg["signers"], docs)
    for s in cfg["signers"]:
        print(f"   ✍️  {s['name']} <{s['email']}> — ordre {s.get('order', 1)}")

    envelope = EnvelopeDefinition(
        email_subject=cfg["subject"],
        email_blurb=cfg.get("message", ""),
        documents=docs,
        recipients=Recipients(signers=signers),
        status="sent",
    )

    envelopes_api = EnvelopesApi(client)
    try:
        result = envelopes_api.create_envelope(account_id=ACCOUNT_ID, envelope_definition=envelope)
        print(f"\n✅ Enveloppe envoyée.")
        print(f"   Envelope ID : {result.envelope_id}")
        print(f"   Statut     : {result.status}")
        print(f"   Dashboard  : https://app{'demo' if DS_ENV == 'demo' else ''}.docusign.com/")
    except ApiException as e:
        sys.exit(f"❌ Erreur DocuSign : {e}")


def list_envelopes(days: int = 30):
    from datetime import datetime, timedelta
    client = _get_api_client()
    api = EnvelopesApi(client)
    from_date = (datetime.utcnow() - timedelta(days=days)).isoformat()
    try:
        result = api.list_status_changes(account_id=ACCOUNT_ID, from_date=from_date)
    except ApiException as e:
        sys.exit(f"❌ Erreur DocuSign : {e}")
    print(f"📋 {len(result.envelopes or [])} enveloppe(s) sur les {days} derniers jours :\n")
    for env in result.envelopes or []:
        print(f"   {env.envelope_id}  |  {env.status:12}  |  {env.email_subject}")


def envelope_status(envelope_id: str):
    client = _get_api_client()
    api = EnvelopesApi(client)
    try:
        env = api.get_envelope(account_id=ACCOUNT_ID, envelope_id=envelope_id)
        recipients = api.list_recipients(account_id=ACCOUNT_ID, envelope_id=envelope_id)
    except ApiException as e:
        sys.exit(f"❌ Erreur DocuSign : {e}")
    print(f"📬 Enveloppe {envelope_id}")
    print(f"   Sujet   : {env.email_subject}")
    print(f"   Statut  : {env.status}")
    print(f"   Créée   : {env.created_date_time}")
    print(f"   Signataires :")
    for signer in recipients.signers or []:
        print(f"      • {signer.name:30}  {signer.status:15}  (ordre {signer.routing_order})")


def download_envelope(envelope_id: str):
    SIGNED_DIR.mkdir(exist_ok=True)
    client = _get_api_client()
    api = EnvelopesApi(client)
    try:
        # Document combiné signé + certificat d'audit
        combined = api.get_document(account_id=ACCOUNT_ID, envelope_id=envelope_id, document_id="combined")
        out_path = SIGNED_DIR / f"{envelope_id}_combined.pdf"
        with open(out_path, "wb") as f:
            f.write(combined)
        print(f"✅ Document combiné signé : {out_path}")

        certificate = api.get_document(account_id=ACCOUNT_ID, envelope_id=envelope_id, document_id="certificate")
        cert_path = SIGNED_DIR / f"{envelope_id}_audit.pdf"
        with open(cert_path, "wb") as f:
            f.write(certificate)
        print(f"✅ Certificat d'audit        : {cert_path}")
    except ApiException as e:
        sys.exit(f"❌ Erreur DocuSign : {e}")


def main():
    parser = argparse.ArgumentParser(description="DocuSign — Envoi des documents de gouvernance")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--envelope", help="Clé de l'enveloppe à envoyer (voir envelopes.json)")
    group.add_argument("--list", action="store_true", help="Liste les enveloppes récentes")
    group.add_argument("--status", help="Vérifie le statut d'une enveloppe (ID)")
    group.add_argument("--download", help="Télécharge les documents signés d'une enveloppe (ID)")
    args = parser.parse_args()

    if args.envelope:
        send_envelope(args.envelope)
    elif args.list:
        list_envelopes()
    elif args.status:
        envelope_status(args.status)
    elif args.download:
        download_envelope(args.download)


if __name__ == "__main__":
    main()
