# Consolidation Fonds Non Cotés — Althos

Trois outils pour le même besoin : compléter la consolidation d'un client avec, pour chaque fonds
non coté détenu :

- la date du **prochain ordre d'entrée** (souscription) possible,
- la date du **prochain ordre de sortie** (rachat) possible,
- la date de la **prochaine réception du cash** suite à un rachat,
- si le **client est concerné par une pénalité de sortie** au vu de sa date d'investissement.

- **`ajout-calendrier-sortie.html`** ⭐ — le conseiller **dépose son fichier de consolidation
  client réel** (glisser-déposer) ; la feuille « Calendrier de sortie » est générée et
  téléchargée automatiquement, sans aucune saisie. C'est l'outil à utiliser au quotidien —
  voir la section dédiée ci-dessous.
- **`index.html`** — outil de saisie manuelle fonds par fonds, pour un usage ponctuel sans
  partir d'un fichier Excel existant (ex. réponse rapide au téléphone).
- **`scripts/build_client_workbook.py`** — équivalent en ligne de commande (Python) du premier
  outil, pour un traitement par script / en lot.

## Utilisation (dépôt automatisé — `ajout-calendrier-sortie.html`)

Ouvrir `ajout-calendrier-sortie.html` dans un navigateur (double-clic, aucun serveur requis,
aucune donnée envoyée sur internet — tout se passe dans le navigateur).

1. Glisser-déposer (ou cliquer pour parcourir) le fichier Excel de consolidation du client.
2. Le traitement est automatique : une feuille **« Calendrier de sortie »** est ajoutée juste
   après la feuille « Consolidation », avec les mêmes lignes de fonds et la même mise en forme.
   Seules les colonnes « Contrat n / TOTAL / Risque / SRI / VL / Mouvements » sont remplacées par
   les 5 colonnes utiles : `Date d'investissement` (à saisir), `Prochain ordre — entrée`,
   `Prochain ordre — sortie`, `Prochaine réception du cash`, `Pénalité de sortie`.
3. Le fichier complété se télécharge automatiquement (`nomdufichier - avec calendrier de sortie.xlsx`).
   Un bouton permet de le retélécharger si besoin.
4. Ouvrir le fichier téléchargé, saisir la date d'investissement pour chaque fonds non coté
   détenu par le client : les colonnes se recalculent automatiquement (formules Excel).

**Comment ça détecte les lignes ?** Aucune structure de fichier n'est supposée fixe : l'outil
repère la feuille via son en-tête (« Support » en colonne A), puis distingue les lignes de
catégorie (bandeau en gras/coloré) des lignes de fonds, et reconnaît un fonds non coté par son
ISIN (comparé à la base Althos) — peu importe le nombre de contrats, l'ordre des fonds ou le
nombre de lignes du fichier déposé. Les fonds cotés (Actions, Fonds prudents…) ou non reconnus
affichent simplement « — ». Testé avec succès sur deux fichiers de consolidation clients réels
aux structures différentes (17 vs 9 contrats, 483 vs 294 lignes).

### Comment ça marche techniquement

Tout tourne dans le navigateur via [ExcelJS](https://github.com/exceljs/exceljs) (vendorisé dans
`vendor/exceljs.min.js`, aucun accès réseau nécessaire) — voir `js/exit_calendar_builder.js`.
C'est un portage JavaScript de `scripts/build_client_workbook.py` (mêmes formules, même logique
de calcul — voir la section dédiée à ce script plus bas pour le détail), rendu robuste face à des
fichiers clients réels dont la structure (nombre de contrats, position des catégories, nombre de
lignes) varie d'un dossier à l'autre.

⚠️ Même limite que pour le script Python : le recalcul automatique des formules n'a pas pu être
vérifié par exécution réelle dans cet environnement (LibreOffice indisponible ici — voir plus
bas). Les formules ont été relues manuellement et testées de bout en bout (dépôt → génération →
relecture de la structure et des formules) sur les deux fichiers clients réels fournis, sans
erreur JavaScript ni incohérence de structure détectée. Vérifiez à l'ouverture chez vous qu'aucune
cellule n'affiche `#NAME?` / `#REF!` / `#VALUE!`.

## Utilisation (page web de saisie manuelle — `index.html`)

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
ajout-calendrier-sortie.html        Outil principal, PRÊT À DISTRIBUER (fichier unique, autonome)
index.html                          Outil de saisie manuelle, PRÊT À DISTRIBUER (fichier unique, autonome)
ajout-calendrier-sortie.template.html   Source éditable de l'outil principal (à éditer, PAS le .html)
index.template.html                     Source éditable de l'outil de saisie (à éditer, PAS le .html)
scripts/inline_assets.py            Régénère les 2 .html autonomes à partir des .template.html
js/exit_calendar_builder.js         Logique de génération (portage JS de scripts/build_client_workbook.py)
vendor/exceljs.min.js               Bibliothèque ExcelJS vendorisée (lecture/écriture .xlsx dans le navigateur)
data/funds_data.js                  Données des fonds générées (voir ci-dessous)
scripts/build_data.py               Script qui regénère data/funds_data.js à partir des fichiers source
scripts/build_client_workbook.py    Équivalent Python (ligne de commande) de ajout-calendrier-sortie.html
source/Calendriers_de_fonds_Althos.xlsx     Classeur "Calendriers" (feuilles Suivi / Calendriers)
source/Bibliotheque_de_fonds_Althos.xlsx    Classeur "Bibliothèque de fonds" (feuille Fonds)
source/ConsolidationTemplateAlthosAI_V5.xlsx  Template de consolidation vierge, utilisé par build_client_workbook.py
```

**`index.html` et `ajout-calendrier-sortie.html` sont des fichiers 100% autonomes** (ExcelJS et
les données des fonds sont inlinés dedans, aucun fichier `js/`, `vendor/` ou `data/` requis à
côté) — c'est ce qui permet de les envoyer/ouvrir comme un fichier unique, y compris directement
depuis l'aperçu d'un .zip sans l'extraire au préalable. **Ne les éditez jamais directement** :
modifiez `*.template.html` (ou `js/exit_calendar_builder.js` / `data/funds_data.js`) puis
relancez `python3 scripts/inline_assets.py` pour régénérer les 2 fichiers finaux.

⚠️ **Ne jamais committer de fichier de consolidation client réel dans ce dépôt** (même anonymisé) :
seul le template vierge ci-dessus doit être versionné. Les fichiers clients sont à traiter
localement via `ajout-calendrier-sortie.html`, qui ne transmet rien en dehors du navigateur.

## Mettre à jour les données (nouveau calendrier annuel, nouvelle pénalité, nouveau fonds…)

1. Remplacer le(s) fichier(s) dans `source/` par la version à jour (mêmes noms de fichier).
2. Relancer :
   ```bash
   python3 scripts/build_data.py
   ```
   Cela régénère entièrement `data/funds_data.js`. Le script affiche un résumé (nombre de fonds,
   nombre de calendriers, répartition des règles de pénalité reconnues).
3. Régénérer les 2 pages autonomes avec les données à jour :
   ```bash
   python3 scripts/inline_assets.py
   ```
4. Ouvrir `index.html` et vérifier rapidement 1-2 fonds connus (sondage), comme préconisé dans
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

## Utilisation (classeur Excel du conseiller)

Le conseiller travaille normalement dans son classeur de consolidation habituel (ex.
`ConsolidationTemplateAlthosAI_V5.xlsx`, feuille `Consolidation` : un fonds par ligne, un
contrat par colonne, avec les montants investis). `scripts/build_client_workbook.py` prend ce
classeur et lui ajoute une nouvelle feuille **« Calendrier de sortie »**, positionnée juste après
`Consolidation`, avec **exactement la même mise en forme** (mêmes lignes de fonds, mêmes
catégories, mêmes couleurs) — seules les colonnes changent :

| Colonne | Contenu |
|---|---|
| Support / ISIN & DIC | Identiques à la feuille Consolidation (copiées telles quelles) |
| Date d'investissement | 🟡 à saisir par le conseiller, pour chaque fonds non coté détenu par le client |
| Prochain ordre — entrée | Cut-off + date de valorisation (VL) de la prochaine fenêtre de souscription |
| Prochain ordre — sortie | Cut-off + date de valorisation (VL) de la prochaine fenêtre de rachat |
| Prochaine réception du cash | Date de règlement du prochain rachat |
| Pénalité de sortie | Statut calculé (concerné / non concerné / à vérifier) à partir de la date d'investissement saisie |

Pour les fonds **cotés** (Actions, Fonds prudents, monétaires, obligataires…), ces 4 colonnes
affichent simplement « — » : le calendrier de sortie ne concerne que les fonds non cotés.

### Générer / régénérer la feuille

```bash
python3 scripts/build_client_workbook.py [chemin_consolidation.xlsx] [chemin_sortie.xlsx]
```

Sans arguments, lit `source/ConsolidationTemplateAlthosAI_V5.xlsx` et écrit
`output/ConsolidationTemplateAlthosAI_V5_avec_calendrier.xlsx`. Pour l'utiliser sur le classeur
réel d'un client (avec les montants déjà saisis), passer son chemin en premier argument — la
feuille `Consolidation` du client n'est jamais modifiée, seule la nouvelle feuille est ajoutée.

### Comment ça marche (toutes les colonnes calculées sont des formules Excel)

Le script ajoute aussi deux feuilles de données masquées, alimentées par
`Calendriers_de_fonds_Althos.xlsx` (même logique que `build_data.py`, voir plus haut) :

- **`BDD_Calendrier`** : une ligne par fonds / type (Souscription ou Rachat) / échéance mensuelle.
- **`BDD_Penalites`** : une ligne par fonds, avec la règle de pénalité réduite à au plus 4 paliers
  "seuil en mois → taux" + un taux "au-delà" (`build_tier_columns()`), pour rester calculable par
  une formule `IFS` en cascade sans jamais recourir à une formule matricielle (CSE).

Les 5 colonnes visibles s'appuient sur des colonnes de calcul intermédiaires masquées (mois de
détention, prochaine échéance trouvée via `MINIFS`/`MAXIFS`, palier de pénalité applicable…). Tout
se recalcule automatiquement à l'ouverture du fichier, à la date du jour — rien n'est figé.

**⚠️ Limite connue de cette tâche : le recalcul automatique n'a pas pu être vérifié dans cet
environnement.** LibreOffice (utilisé habituellement pour valider les formules avant livraison)
n'a pas réussi à charger le moindre fichier `.xlsx` dans ce bac à sable — y compris les fichiers
sources originaux, jamais modifiés — ce qui pointe vers un problème d'infrastructure de la
session, indépendant de ce script. Les formules ont donc été relues manuellement avec soin
(cohérence des plages, parenthésage, gestion des cas vides/erreurs) plutôt que vérifiées par
exécution réelle. **À l'ouverture du fichier dans Excel ou LibreOffice Calc chez vous, vérifiez
qu'aucune cellule n'affiche `#NAME?`, `#REF!` ou `#VALUE!`** avant de vous fier au résultat ; le
cas échéant, signalez-le pour correction.
