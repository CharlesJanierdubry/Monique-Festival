# Registres légaux — Monique Festival

*Conformité loi du 1er juillet 1901 — ordonnance n° 2015-904 du 23/07/2015 (registres électroniques autorisés).*

## Pourquoi ces registres

Trois registres ont une valeur juridique pour une association loi 1901 :

| Registre | Obligation | Référence légale |
|---|---|---|
| **Registre spécial** | ✅ Obligatoire | Article 5 loi 1901 |
| **Registre des délibérations du Bureau** | ⚠️ Si prévu par les statuts | Statuts art. 11 et 13 |
| **Registre des membres** | ⚠️ Si prévu par les statuts ou pour traçabilité RGPD | Statuts + bonne pratique |

S'y ajoute le **registre des dons** ([../../officiel/dons/Registre_dons_2026.md](../../officiel/dons/Registre_dons_2026.md)) — non obligatoire mais essentiel pour le rescrit fiscal et les contrôles DDFIP.

## Support choisi : électronique conforme

Conformément à l'**ordonnance n° 2015-904 du 23 juillet 2015**, les registres sont tenus en format électronique et respectent les 5 conditions de validité :

1. **Horodatage certain** — par signature électronique DocuSign à chaque mise à jour majeure
2. **Numérotation chronologique** — entrées numérotées sans rupture
3. **Pas de rature** — uniquement des ajouts datés (les corrections font l'objet d'une nouvelle entrée)
4. **Conservation au siège social** — Drive Google + dépôt local au 54 chemin de Valentin
5. **Lisibilité dans le temps** — format Markdown + PDF signé archivés en parallèle

## Workflow

À chaque modification statutaire, élection, délibération du Bureau ou nouvelle adhésion :

1. **Mettre à jour le fichier `.md`** correspondant (ajout d'entrée numérotée, jamais de modification rétroactive)
2. **Commit Git** avec message clair (horodatage immutable supplémentaire)
3. **Trimestriellement** (ou immédiatement si entrée critique) : exporter le `.md` en PDF, faire signer par la Présidente via DocuSign, archiver le PDF signé dans [signes/](signes/)
4. **Sauvegarde Drive** dans `1_Gouvernance/Registres/`

## Composition

```
gouvernance/registres/
├── README.md                       (ce fichier)
├── Registre_special.md             (modifications statutaires + composition Bureau)
├── Registre_deliberations.md       (index des PV de Bureau et AGE)
├── Registre_membres.md             (liste des membres adhérents)
└── signes/                         (PDF signés DocuSign)
```

## En cas de contrôle (DDFIP, Préfecture, contentieux)

Présenter le **PDF signé DocuSign le plus récent** de chaque registre — la signature électronique avancée eIDAS a la même valeur probante que le registre papier coté/paraphé traditionnel.

Pour la traçabilité fine, l'historique Git complet est consultable depuis `gouvernance/registres/`.
