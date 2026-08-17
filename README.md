# Consolidation Fonds Non Cotés — Althos

Outil interne (page HTML autonome, sans backend) pour consolider les investissements d'un
client dans des fonds non cotés et, à la demande, afficher automatiquement pour chaque ligne :

- la date du **prochain ordre d'entrée** (souscription) possible,
- la date du **prochain ordre de sortie** (rachat) possible,
- la date de la **prochaine réception du cash** suite à un rachat,
- si le **client est concerné par une pénalité de sortie** au vu de sa date d'investissement.

## Utilisation

Ouvrir `index.html` directement dans un navigateur (double-clic, aucun serveur requis).

1. Renseigner le nom du client (facultatif, affiché en en-tête et à l'impression).
2. Cliquer sur **+ Ajouter une ligne** pour chaque fonds détenu par le client, choisir le fonds
   (recherche par nom ou ISIN), la date d'investissement, le montant et la devise.
3. Si le client demande ses dates de sortie, cliquer sur **Afficher les dates de sortie** :
   4 colonnes supplémentaires apparaissent avec les informations calculées à la date du jour.
4. **Imprimer / PDF** génère une version imprimable propre (sans les boutons d'édition) pour le
   dossier client.

Les lignes saisies sont conservées dans le navigateur (`localStorage`) — rien n'est envoyé sur
un serveur. Vider le cache du navigateur ou cliquer sur **Tout effacer** réinitialise l'outil.

⚠️ **Toujours vérifier une donnée sensible (pénalité, date d'ordre) contre le PDF officiel du
fonds avant de répondre à un client** — voir le bandeau d'avertissement affiché en haut de la
page. L'outil est une aide à la consolidation, pas une source réglementaire.

## Structure du dépôt

```
index.html                          Page de l'outil (HTML + CSS + JS, aucune dépendance externe)
data/funds_data.js                  Données des fonds générées (voir ci-dessous) — chargées par index.html
scripts/build_data.py               Script qui regénère data/funds_data.js à partir des fichiers source
source/Calendriers_de_fonds_Althos.xlsx     Classeur "Calendriers" (feuilles Suivi / Calendriers)
source/Bibliotheque_de_fonds_Althos.xlsx    Classeur "Bibliothèque de fonds" (feuille Fonds)
```

## Mettre à jour les données (nouveau calendrier annuel, nouvelle pénalité, nouveau fonds…)

1. Remplacer le(s) fichier(s) dans `source/` par la version à jour (mêmes noms de fichier).
2. Relancer :
   ```bash
   python3 scripts/build_data.py
   ```
   Cela régénère entièrement `data/funds_data.js`. Le script affiche un résumé (nombre de fonds,
   nombre de calendriers, répartition des règles de pénalité reconnues).
3. Ouvrir `index.html` et vérifier rapidement 1-2 fonds connus (sondage), comme préconisé dans
   la feuille « Lisez-moi » du classeur Calendriers.
4. Committer les 3 fichiers modifiés (`source/*.xlsx`, `data/funds_data.js`).

### D'où viennent les données affichées

- **Liste des fonds non cotés + pénalités de sortie** : feuille `Suivi` du classeur
  Calendriers (complétée par les fonds supplémentaires trouvés dans `Calendriers par fonds`,
  ex. Antheor).
- **Calendrier mensuel (cut-off, VL, exécution, publication VL, règlement/réception cash)** :
  feuille `Calendriers` du classeur Calendriers — une ligne par fonds / année / type
  (Souscription ou Rachat) / mois. Seuls les fonds ayant un calendrier officiel déposé y
  figurent (37 fonds au moment de la première génération, sur 144 fonds non cotés suivis) ;
  pour les autres, l'outil affiche « Calendrier non disponible ».
- **Devise, SRI, frais, liquidité, temporalité, lien DICI** : feuille `Fonds` du classeur
  Bibliothèque de fonds, rattachée par ISIN.

### Comment les pénalités de sortie sont interprétées

Le texte libre de la colonne « Lock-up / pénalité de sortie » est transformé en règle
structurée par `parse_penalite()` dans `scripts/build_data.py` :

| Type détecté | Exemple de texte source | Comportement dans l'outil |
|---|---|---|
| `aucune` | « Aucune pénalité actuellement. » / « Pas de lock-up. » | Affiché en vert, pas de calcul nécessaire |
| `seuil` | « Pénalité de 5% ... avant 1 an de détention » | Comparé à la date d'investissement saisie : concerné (rouge) tant que la détention est inférieure au seuil, sinon non concerné (vert) |
| `soft` | « Soft lock-up de 2% si rachat dans les 12 mois... » | Même calcul que `seuil`, libellé « soft lock-up » ajouté |
| `degressif` | « Pénalité dégressive : 0-18 mois 7,5% · 18-36 mois 5% · ... » | Palier applicable déterminé selon la détention ; 0% = non concerné |
| `manuel` | Formulations ambiguës ou taux multiples selon la part détenue (ex. Hg Fusion) | **Aucun calcul automatique** — le texte brut est affiché avec un avertissement « à vérifier manuellement » |
| `inconnue` | Pas de texte renseigné dans la base | « Pénalité non renseignée — vérifier la notice / DICI » |

Le texte brut d'origine est **toujours affiché** sous le statut calculé, pour permettre une
vérification en un coup d'œil avant de répondre à un client — conformément à la procédure de
vérification par sondage décrite dans la feuille « Lisez-moi » du classeur Calendriers.

### Comment les prochaines dates d'ordre sont calculées

Pour un fonds donné et une direction (entrée = Souscription, sortie = Rachat), l'outil parcourt
les échéances mensuelles du calendrier (triées par date), et retient la première dont la date
de cut-off est encore dans le futur au moment de la consultation (date du navigateur de
l'utilisateur, recalculée à chaque ouverture — jamais figée). Si le calendrier connu ne
couvre plus la période courante (ex. calendrier 2025 non encore renouvelé pour 2026), un
message invite à demander le calendrier à jour plutôt que d'afficher une date erronée.
