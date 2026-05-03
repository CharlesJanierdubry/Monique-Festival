# Mail type — envoi d'un reçu fiscal aux donateurs rétroactifs

*Template à utiliser pour envoyer les 14 reçus fiscaux rétroactifs (dons du 20-23 avril 2026).*

---

## Objet

```
Monique Festival — Votre reçu fiscal pour votre don du {{DATE_DON}}
```

## Corps du mail

```
Bonjour {{PRENOM}},

Je vous écris au nom de l'association Monique Festival — vous nous avez
soutenus le {{DATE_DON}} en versant un don de {{MONTANT}} € sur notre
campagne de crowdfunding.

Un immense merci : ce premier coup de pouce nous aide à concrétiser
la première édition du Monique Festival, qui aura lieu à La Grange
Huguenet (Besançon) les 28, 29 et 30 août 2026.

Vous trouverez en pièce jointe votre reçu fiscal (Cerfa n° 11580*04).
Il vous permet de bénéficier d'une réduction d'impôt sur le revenu
égale à {{TAUX_REDUCTION}} % du montant de votre don, à reporter
dans votre déclaration 2027 pour les revenus 2026.

Petite précision administrative : notre association s'appelait
initialement JD Production et change de dénomination pour Monique
Festival suite à l'Assemblée Générale du 19 avril 2026. Les reçus
fiscaux sont désormais émis sous le nom Monique Festival. La personne
morale, elle, est inchangée — mêmes numéros RNA et SIREN.

D'ici là, vous pouvez nous suivre sur Instagram : @monique.festival
— la programmation se dévoile au rythme de 3 artistes par semaine.

À très vite, et encore merci !

Judith Laithier
Présidente de l'association Monique Festival

www.monique-festival.fr (bientôt)
info@monique-festival.fr

---
PJ : Reçu fiscal n° {{NUMERO_RECU}}
```

## Placeholders à remplacer

| Placeholder | Exemple |
|---|---|
| `{{PRENOM}}` | Lise |
| `{{DATE_DON}}` | 23/04/2026 |
| `{{MONTANT}}` | 30 |
| `{{TAUX_REDUCTION}}` | 66 (particuliers) ou 60 (entreprises) |
| `{{NUMERO_RECU}}` | 15052026-MF-000001 |

## Notes d'usage

- Un mail par donateur (14 en tout pour les dons du 20-23/04)
- Envoi recommandé **après** mise à jour Hello Asso (post-récépissé préfectoral)
- Pièce jointe : le PDF du reçu fiscal généré par le script `generate_recus_fiscaux.py`
- **Ton** : chaleureux mais sobre, à adapter si tu connais bien le donateur
- Une liste BCC à 14 serait maladroite — envois individuels recommandés pour préserver la personnalisation

## Version courte *(pour les dons < 50 €, ton plus léger)*

```
Bonjour {{PRENOM}},

Merci encore pour votre don de {{MONTANT}} € à Monique Festival !
Vous trouverez en pièce jointe votre reçu fiscal qui vous donnera
droit à une réduction d'impôt de {{TAUX_REDUCTION}} % lors de votre
prochaine déclaration.

L'association, précédemment JD Production, s'appelle maintenant
Monique Festival (même personne morale, même SIREN).

À bientôt, sur Instagram @monique.festival ou à La Grange Huguenet
les 28-30 août !

Judith Laithier — Présidente

PJ : Reçu fiscal n° {{NUMERO_RECU}}
```
