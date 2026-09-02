# Consolidation Fonds Non Cotés — Althos

Trois outils pour le même besoin : compléter la consolidation d'un client avec, pour chaque fonds
non coté **détenu par ce client et disposant d'un calendrier de rachat (sortie) connu** :

- les 5 dates du **calendrier de rachat** telles qu'enregistrées dans la base Althos : ordre
  avant (cut-off), VL, exécuté, publié, cash reçu,
- si le **client est concerné par une pénalité de sortie** au vu de sa date d'investissement (la
  cellule reste **vide** dès que le délai de pénalité est dépassé — un message n'apparaît que
  s'il y a quelque chose à signaler : concerné, aucune pénalité prévue, à vérifier manuellement,
  ou non renseignée).

Il n'y a volontairement **aucune colonne pour les ordres d'entrée (souscription)** : seul le
rachat (sortie) est traité.

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
   après la feuille « Consolidation ». C'est un tableau **compact** listant uniquement les fonds
   que ce client **détient réellement** (montant total non nul dans sa consolidation) **et** pour
   lesquels un **calendrier de rachat (sortie) officiel est connu** dans la base Althos — pas la
   liste complète des fonds de la consolidation, pas les fonds cotés, pas les fonds non détenus,
   pas les fonds dont seul le calendrier de souscription est connu. Pour chaque fonds retenu :
   `Fonds`, `ISIN`, `Date d'investissement` (à saisir), `Rachat — ordre avant`, `Rachat — VL`,
   `Rachat — exécuté`, `Rachat — publié`, `Rachat — cash reçu` (les 5 dates telles quelles dans la
   base, pour l'échéance de rachat la plus proche à partir d'aujourd'hui), `Pénalité de sortie`
   (vide si le délai de pénalité est dépassé, sinon un message explicite, sans icône). Si aucun
   fonds détenu n'a de calendrier de rachat connu, la feuille l'indique clairement au lieu d'un
   tableau vide.

   **Un bandeau par titulaire, suivi de sa propre ligne d'en-tête** : quand la Consolidation
   subdivise les contrats par titulaire (ligne juste au-dessus de l'en-tête « Support », ex.
   « Monsieur » / « Madame » / un prénom / une société), chaque titulaire obtient son propre
   bandeau (harmonisé en « M. Prénom NOM » / « Mme Prénom NOM » à partir du titre de Consolidation,
   qui énonce explicitement le genre — jamais deviné ; une société garde son nom tel quel), suivi
   immédiatement de la ligne d'en-tête (`Fonds`, `ISIN`, ...) répétée, puis de ses catégories et
   ses fonds — plutôt qu'une colonne « Titulaire » mélangeant tout le monde. Un même fonds détenu
   à la fois par Monsieur ET Madame apparaît donc dans les 2 sections, chacune avec sa propre date
   d'investissement et son propre statut de pénalité. Une ligne vide (non bordée) sépare 2
   titulaires consécutifs, pour que leurs tableaux ne paraissent pas collés — le contour
   d'impression englobe malgré tout l'ensemble des titulaires. Une seule ligne d'en-tête, sans
   bandeau, si le fichier ne subdivise pas les contrats (client à titulaire unique).

   La présentation reprend celle de la Consolidation elle-même : un bandeau de titre centré
   « Calendrier des délais de sortie » (même bleu, même police que le titre de Consolidation),
   un sous-titre daté en italique juste en dessous (date du jour au format jj/mm/aaaa, même style
   que le sous-titre de Consolidation), un bandeau par titulaire dans le même bleu que l'en-tête
   du tableau (pas la police, parfois plus petite, de la ligne où Consolidation affiche déjà
   « Monsieur »/« Madame » au-dessus de ses colonnes contrat — ce serait un bandeau pleine largeur
   à une taille de sous-étiquette), un en-tête de tableau dans ce même bleu, des bandeaux de
   catégorie dans la couleur des catégories de Consolidation, les fonds sur fond blanc, et un
   quadrillage dont la couleur est reprise telle quelle de la bordure déjà utilisée par
   Consolidation (pas une couleur codée en dur : elle dépend du thème propre à chaque classeur,
   donc peut être beige, bleue, ou autre selon le fichier). Le quadrillage Excel par défaut est
   désactivé, comme sur Consolidation, pour un fond bien blanc. Police uniforme sur toute la
   feuille (celle utilisée par la Consolidation, pas un choix arbitraire).

   **Colonne Pénalité de sortie** avec un retour à la ligne explicite entre chaque phrase du
   message (ex. « Pénalité de sortie : ... » puis, sur sa propre ligne, « Durée de détention
   actuelle : X mois. »). Chaque ligne de fonds a une hauteur de 75pt par défaut — augmentée à
   partir de la longueur réelle du texte de la base pour ce fonds précis, connue à la génération
   (jamais un ajustement automatique par Excel : peu fiable pour du texte arrivé par formule
   plutôt que saisi à la main). L'estimation du nombre de lignes nécessaires applique une marge de
   sécurité volontairement large (police Montserrat, plus large qu'une police système classique,
   et cet environnement de génération ne dispose d'aucun moteur de rendu pour vérifier le résultat
   pixel par pixel) — mieux vaut une ligne visiblement plus haute que nécessaire qu'un texte coupé.

   La zone d'impression est définie sur tout le tableau généré, en orientation paysage et mise à
   l'échelle sur une page en largeur (comme sur Consolidation) — imprimer ou exporter en PDF
   depuis cette feuille donne donc directement un rendu propre, sans réglage manuel. La feuille
   s'ouvre aussi en vue « aperçu des sauts de page » (comme Consolidation), ce qui affiche
   automatiquement le contour bleu de la zone d'impression dès l'ouverture, sans manipulation.

   **Fonds fermés** : certains fonds n'acceptent plus aucun rachat (hors cas exceptionnel prévu
   au règlement, ex. décès, invalidité) — la base Althos les marque « fond fermé ». Ces fonds sont
   retenus dans la feuille **même sans calendrier de rachat**, avec le message « Fonds fermé :
   aucun rachat possible. » suivi, quand la base le précise (feuille « Calendriers par fonds »),
   de la durée de vie du fonds et de ses conditions de prorogation (segment « Durée de vie du
   fonds ») **et**, séparément, de la durée pendant laquelle l'investisseur est bloqué sans
   possibilité de sortie (segment « Conservation obligatoire » ou une de ses variantes de
   libellé dans la base — « blocage / conservation », « durée initiale »...), les 2 informations
   étant distinctes et pouvant coexister sur 2 lignes (durée de vie du fonds ; durée de blocage
   effective de l'investisseur), pour que le conseiller ne les découvre pas au moment où le
   client demande à sortir.

   **La feuille Consolidation elle-même est aussi complétée**, avec 2 colonnes `Rachat — cash
   reçu` et `Pénalité de sortie`, en formule vers la valeur déjà calculée dans « Calendrier de
   sortie » pour le même ISIN — **la même information dans les 2 pages**, jamais recalculée
   séparément — pour que le conseiller ait l'essentiel sous les yeux sans changer d'onglet. Elles
   sont ajoutées collées juste après la dernière colonne réellement utile du tableau (ex.
   « Dernière Valeur Liquidative ») ; si une colonne « Mouvements en cours » existe déjà tout à la
   fin du tableau (toujours en dernière position), elle est repoussée de 2 colonnes vides après
   les 2 nouvelles, plutôt que les nouvelles colonnes n'atterrissent après elle. Leur en-tête
   occupe 2 lignes fusionnées comme leurs voisines (« Dernière Valeur Liquidative », etc.) quand
   la Consolidation utilise elle-même cette mise en page, pour un bleu marine bien aligné en
   hauteur. Rempli uniquement pour les fonds réellement détenus (même filtre « montant non nul »
   que « Calendrier de sortie »). Si un même fonds est détenu via plusieurs titulaires (donc
   plusieurs lignes dans « Calendrier de sortie »), la ligne unique de Consolidation reprend la
   première trouvée ; le détail par titulaire reste dans l'autre feuille. Colonne Pénalité de
   sortie de largeur 30 (plus étroite que sur « Calendrier de sortie », pour rester proportionnée
   aux colonnes voisines de Consolidation, plus compactes), quadrillage complet fin (même couleur
   reprise de Consolidation) — la hauteur de ligne d'origine du fichier client est conservée SAUF
   si le texte de pénalité pour ce fonds précis (durée de vie ET durée de blocage pour un fonds
   fermé comprises — les 2 sont prises en compte dans le calcul, pas seulement la 1re) a besoin de
   plus de place (calculé comme sur « Calendrier de sortie », voir plus haut, avec la même
   largeur de colonne réduite) : le texte lisible passe avant la régularité visuelle des lignes,
   mais la grande majorité des lignes (courtes ou vides) garde sa hauteur native, seules les
   quelques lignes à texte long sont agrandies — au besoin très nettement, plutôt que de risquer
   un texte coupé. Le
   quadrillage court sans interruption sur toute la hauteur du VRAI tableau de fonds (catégories
   et fonds, détenus ou non par ce client — Consolidation liste l'univers complet des fonds), sans
   déborder sur un éventuel petit tableau récapitulatif hors-tableau plus bas sur la feuille (ex.
   « dont actions européennes : 15 % ») : une ligne n'y est traitée comme une vraie ligne de fonds
   que si sa colonne TOTAL porte une vraie agrégation (=SUM(...) sur ses colonnes de contrat), pas
   un simple renvoi vers une autre cellule. Le fichier d'origine ne ferme pas le tableau par une
   bordure basse sur sa dernière ligne, mais par une bordure haute sur la ligne suivante (souvent
   une note en italique) — sur les colonnes d'origine uniquement : la présence (mais pas
   l'épaisseur) de cette bordure de clôture est donc détectée et reproduite sur les 2 nouvelles
   colonnes à cette même ligne, avec le même trait fin que le reste de leur quadrillage (jamais le
   cadre épais d'origine), pour que le trait de fermeture du tableau ne s'arrête pas juste avant
   elles tout en restant cohérent avec le reste des 2 colonnes.
   Le bandeau de titre tout en haut est étendu pour
   couvrir les 2 nouvelles colonnes. Consolidation utilise 2 teintes de beige distinctes selon le
   niveau d'imbrication d'une catégorie (ex. catégorie principale en beige foncé, sous-catégorie
   en beige clair) : le fond de chaque bandeau de catégorie (sur les 2 nouvelles colonnes comme
   sur « Calendrier de sortie ») reprend la teinte exacte de SA catégorie d'origine, pas un modèle
   unique emprunté à la première catégorie trouvée. Une ligne de catégorie (bandeau) est détectée
   par son fond uni ET l'absence d'ISIN en colonne B — jamais par le gras seul, que certains
   fichiers clients réservent aux catégories de premier niveau sans l'appliquer aux
   sous-catégories (ex. « dont actions européennes »), qui restent pourtant de vrais bandeaux à
   traiter de la même façon. L'épaisseur et la couleur du quadrillage sont toutes deux reprises
   telles quelles de la bordure fine déjà utilisée par Consolidation (jamais une épaisseur forcée
   en dur), pour un rendu identique. Les dates de la colonne « Rachat — cash reçu »
   et le texte de la colonne « Pénalité de sortie » sont en gras (sur les deux feuilles) — une
   formule Excel ne pouvant jamais renvoyer un texte à mise en forme mixte (gras partiel sur une
   phrase précise au milieu du message), c'est tout le contenu de la cellule qui est mis en gras
   en bloc.
3. **Si des fonds détenus ont une pénalité de sortie qui dépend d'une date d'investissement**
   (règle à seuil, soft lock-up ou dégressive), la page affiche directement une liste de ces
   fonds avec un champ de saisie par fonds — le conseiller renseigne les dates connues sans
   ouvrir Excel, ce qui lui montre immédiatement quels fonds sont concernés (les fonds sans
   pénalité datée, fermés, ou à vérifier manuellement n'apparaissent pas dans cette liste,
   puisqu'ils ne dépendent pas d'une date). Une case laissée vide reste à compléter plus tard,
   directement dans le fichier Excel généré. Cette saisie reste 100% locale au navigateur — rien
   n'est envoyé sur un serveur, il n'y a pas d'intranet à héberger. Un bouton permet de générer le
   fichier sans saisir de date si besoin.
4. Le fichier complété se télécharge automatiquement (`nomdufichier - avec calendrier de sortie.xlsx`).
   Un bouton permet de le retélécharger si besoin.
5. Pour les dates non saisies à l'étape 3, ouvrir le fichier téléchargé et compléter la date
   d'investissement pour le fonds concerné : les 5 dates de rachat et la pénalité de sortie se
   recalculent automatiquement (formules Excel), à partir des échéances les plus proches **de la
   date du jour** au moment de l'ouverture (ou de la date de retrait simulée si l'étape 0
   ci-dessous a été utilisée). Si un fonds a bien des échéances de rachat connues dans la base mais
   qu'aucune n'est encore à venir par rapport à cette date (calendrier officiel qui ne couvre pas
   assez loin dans le temps — typique d'une simulation à une date très éloignée, ou simplement un
   fonds dont le calendrier de l'année en cours n'est pas encore publié), les 5 colonnes de rachat
   reprennent automatiquement la **dernière échéance connue** (le calendrier de l'année précédente
   la plus récente disponible dans la base) plutôt que de rester vides ou d'afficher une date
   invalide (« 00/01/1900 ») — avec un **commentaire Excel** (icône rouge, au survol) sur la
   cellule « Rachat — ordre avant » précisant explicitement l'année d'où ces dates sont reprises
   (« calendrier pris sur l'année 2025 », ou l'année correspondante), pour que le conseiller sache
   que ces dates sont une estimation basée sur l'an dernier, pas le calendrier officiel de l'année
   demandée.

**Simuler une date de retrait future (optionnel)** : avant de déposer le fichier, le conseiller
peut renseigner un champ « Simuler une date de retrait future » sur la page. Si le client demande
son retrait dans plusieurs mois plutôt qu'aujourd'hui, cette date remplace alors la date du jour
dans **tous** les calculs du classeur généré (les 2 pages) — échéances de rachat les plus proches,
durée de détention, pénalité de sortie active ou non — comme si on se trouvait déjà à cette date.
Techniquement, une seule cellule de référence (cachée, colonne d'aide de « Calendrier de sortie »)
porte soit `=TODAY()` (comportement par défaut, toujours à jour à l'ouverture), soit la date
choisie (figée) ; toutes les formules concernées la référencent plutôt que d'appeler `TODAY()`
directement. Quand une simulation est active, un bandeau « Si Date de Retrait Exécuté : JJ/MM/AAAA »
apparaît sur les **2 pages** — sous l'en-tête de chaque tableau (par titulaire) sur « Calendrier de
sortie », et juste au-dessus des colonnes « Rachat — cash reçu » / « Pénalité de sortie » sur
Consolidation (même ligne que les libellés de regroupement par contrat déjà présents à cet
endroit), avec la même bordure fine que le reste du quadrillage plutôt que le cadre épais hérité
de l'en-tête (le style d'en-tête est repris pour la couleur/police, mais sa bordure est
explicitement réinitialisée) — pour que ce soit toujours visible, où que le conseiller regarde,
sans casser la régularité visuelle du quadrillage autour. Ce bandeau construit
« JJ/MM/AAAA » lui-même via `DAY()`/`MONTH()`/`YEAR()` plutôt que `TEXT(date,"dd/mm/yyyy")` : les
codes de format de date de `TEXT()` sont traduits selon la langue d'Excel (« jj/mm/aaaa » en
français), donc `"dd/mm/yyyy"` y produit `#VALEUR!` — la construction manuelle, elle, ne dépend
d'aucune langue. Laisser le champ vide
**ou choisir explicitement la date du jour elle-même** revient
exactement au comportement par défaut (formule `=TODAY()` toujours à jour, aucun bandeau) : ce
n'est une simulation que si la date choisie diffère réellement d'aujourd'hui. Côté script Python
(`build_client_workbook.py`), le 3ᵉ argument optionnel en ligne de commande (format `AAAA-MM-JJ`)
fait la même chose.

**Comment ça détecte les fonds à retenir ?** Aucune structure de fichier n'est supposée fixe :
l'outil repère la feuille via son en-tête (« Support » en colonne A), la colonne « TOTAL » (montant
détenu, tous contrats confondus), puis distingue les lignes de catégorie (bandeau en gras/coloré)
des lignes de fonds. Un fonds est retenu si son ISIN correspond à un fonds pour lequel un calendrier
de **rachat** existe dans la base Althos **et** que son montant total détenu est non nul — peu
importe le nombre de contrats, l'ordre des fonds ou le nombre de lignes du fichier déposé. Testé
avec succès sur deux fichiers de consolidation clients réels aux structures différentes (17 vs 9
contrats, 483 vs 294 lignes) : 5 lignes retenues sur l'un (titulaire unique), 13 sur l'autre
(9 fonds, dont certains détenus par plusieurs titulaires à la fois).

### Comment ça marche techniquement

Tout tourne dans le navigateur via [ExcelJS](https://github.com/exceljs/exceljs) (vendorisé dans
`vendor/exceljs.min.js`, aucun accès réseau nécessaire) — voir `js/exit_calendar_builder.js`.
C'est un portage JavaScript de `scripts/build_client_workbook.py` (mêmes formules, même logique
de calcul — voir la section dédiée à ce script plus bas pour le détail), rendu robuste face à des
fichiers clients réels dont la structure (nombre de contrats, position des catégories, nombre de
lignes) varie d'un dossier à l'autre.

⚠️ Même limite que pour le script Python : le recalcul automatique des formules n'a pas pu être
vérifié par exécution réelle d'Excel dans cet environnement (LibreOffice indisponible ici — voir
plus bas). Les formules ont été relues manuellement, testées de bout en bout (dépôt → génération →
relecture de la structure et des formules) sur les deux fichiers clients réels fournis, et
vérifiées avec un moteur de calcul de formules indépendant ([HyperFormula](https://hyperformula.handsontable.com/))
qui les exécute réellement à partir des données du fichier généré — sans erreur ni incohérence
détectée. Le fichier généré ne contenant pas de `calcChain.xml` (ni ExcelJS ni openpyxl n'exécutent
eux-mêmes les formules), l'indicateur `fullCalcOnLoad` est explicitement activé dans le classeur
pour qu'Excel recalcule tout à l'ouverture plutôt que de se fier à un éventuel calcul partiel.
Vérifiez malgré tout à l'ouverture chez vous qu'aucune cellule n'affiche `#NAME?` / `#REF!` /
`#VALUE!` (ou ne reste vide) — au besoin un `Ctrl+Alt+F9` force un recalcul complet.

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
  (Souscription ou Rachat) / mois. **Complété par les mini-tableaux calendrier insérés sous
  certains fonds dans la feuille `Calendriers par fonds`** (ligne d'en-tête « Mois » suivie des
  échéances) : c'est parfois la SEULE source de calendrier pour un fonds (ex. Schroder Semi
  Liquid GPE, absent de la feuille `Calendriers` à plat) — voir `load_calendrier_par_fonds()` et
  `merge_calendriers()` dans `scripts/build_data.py`. En cas de chevauchement, la feuille
  `Calendriers` (entretenue à la main) reste prioritaire ; le second tableau ne comble que les
  trous (nouveaux fonds, mois manquants). 47 fonds couverts au total au moment de cette mise à
  jour (sur 144 fonds non cotés suivis) ; pour les autres, l'outil affiche « Calendrier non
  disponible » (sauf les fonds fermés, retenus quand même — voir plus haut).
- **Devise, SRI, frais, liquidité, temporalité, lien DICI** : feuille `Fonds` du classeur
  Bibliothèque de fonds, rattachée par ISIN.

### Comment les pénalités de sortie sont interprétées

Le texte libre de la colonne « Lock-up / pénalité de sortie » est transformé en règle
structurée par `parse_penalite()` dans `scripts/build_data.py` :

| Type détecté | Exemple de texte source | Comportement dans l'outil |
|---|---|---|
| `aucune` | « Aucune pénalité actuellement. » / « Pas de lock-up. » | Cellule **vide** — pas de pénalité, rien à signaler |
| `seuil` | « Pénalité de 5% ... avant 1 an de détention » | Comparé à la date d'investissement saisie : message « Pénalité de sortie : <texte brut>. Durée de détention actuelle : X mois. » tant que la détention est inférieure au seuil, sinon vide (non concerné) |
| `soft` | « Soft lock-up de 2% si rachat dans les 12 mois... » | Même calcul que `seuil` |
| `degressif` | « Pénalité dégressive : 0-18 mois 7,5% · 18-36 mois 5% · ... » | Palier applicable déterminé selon la détention ; 0% = vide (non concerné) |
| `manuel` | Formulations ambiguës ou taux multiples selon la part détenue (ex. Hg Fusion) | **Aucun calcul automatique** — le texte brut est affiché avec un avertissement « à vérifier manuellement » |
| `inconnue` | Pas de texte renseigné dans la base | Cellule **vide** — pas de pénalité renseignée dans la base équivaut à pas de pénalité |
| `ferme` | « fond fermé » (détecté en priorité sur toute autre mention dans le même texte) | Fonds retenu même sans calendrier de rachat, avec le message « Fonds fermé : aucun rachat possible. » suivi (si connue) de la durée de vie du fonds et de ses conditions de prorogation |

Seuls les fonds `seuil` / `soft` / `degressif` ont une pénalité qui dépend de la date
d'investissement : ce sont eux, et eux seuls, que la page de dépôt automatisé (voir plus haut)
propose de renseigner directement dans le navigateur avant de générer le fichier.

Le texte brut d'origine (quand il existe) est repris tel quel dans le message affiché au
conseiller, pour permettre une vérification en un coup d'œil avant de répondre à un client —
conformément à la procédure de vérification par sondage décrite dans la feuille « Lisez-moi » du
classeur Calendriers.

**La cellule « Pénalité de sortie » reste vide dès qu'il n'y a rien d'actionnable à signaler** :
pas de pénalité prévue pour ce fonds (`aucune`), pénalité non renseignée dans la base (`inconnue`,
traité comme « pas de pénalité »), ou délai de pénalité dépassé (`seuil` / `soft` / `degressif`
avec une détention supérieure au seuil). Un message n'apparaît que dans les cas où le conseiller
doit agir : une règle ambiguë à vérifier à la main (`manuel`), une pénalité active en cours, ou
une saisie manquante/à corriger (date d'investissement pas encore saisie, ou postérieure à
aujourd'hui).

### Comment les prochaines dates de rachat sont calculées

Pour un fonds donné, l'outil parcourt les échéances mensuelles de type **Rachat** du calendrier,
et retient la première dont la date de cut-off (« ordre avant ») est encore dans le futur **au
moment de la consultation** (date du jour, recalculée à chaque ouverture du fichier — jamais
figée). Les 5 dates (ordre avant, VL, exécuté, publié, cash reçu) affichées sont celles de cette
échéance, reprises telles quelles depuis la base — aucune date composée ou recalculée. Si le
calendrier connu ne couvre plus la période courante (ex. calendrier 2025 non encore renouvelé
pour 2026), les colonnes restent vides plutôt que d'afficher une date erronée.

## Utilisation (classeur Excel du conseiller — équivalent en ligne de commande)

Le conseiller travaille normalement dans son classeur de consolidation habituel (ex.
`ConsolidationTemplateAlthosAI_V5.xlsx`, feuille `Consolidation` : un fonds par ligne, un
contrat par colonne, avec les montants investis, et une colonne `TOTAL`). `scripts/build_client_workbook.py`
prend ce classeur et lui ajoute une nouvelle feuille **« Calendrier de sortie »**, positionnée
juste après `Consolidation` — même logique que `ajout-calendrier-sortie.html` (voir plus haut) :
un tableau **compact**, listant uniquement les fonds que ce client **détient réellement**
(colonne `TOTAL` non nulle) **et** pour lesquels un **calendrier de rachat (sortie) est connu**.
Ni les fonds cotés, ni les fonds non détenus, ni les fonds non cotés sans calendrier de rachat
n'apparaissent, et il n'y a pas de colonnes pour les ordres d'entrée — ce n'est pas une copie de
la feuille `Consolidation`.

| Colonne | Contenu |
|---|---|
| Fonds / ISIN | Nom et ISIN du fonds retenu |
| Date d'investissement | 🟡 à saisir par le conseiller |
| Rachat — ordre avant | Date de cut-off de la prochaine échéance de rachat (à partir d'aujourd'hui) |
| Rachat — VL | Date de valorisation associée |
| Rachat — exécuté | Date d'exécution de l'ordre |
| Rachat — publié | Date de publication de la VL |
| Rachat — cash reçu | Date de règlement / réception du cash |
| Pénalité de sortie | Vide dès qu'il n'y a rien à signaler (aucune pénalité prévue, non renseignée dans la base, ou délai dépassé) ; message si concerné par une pénalité active, si la règle est ambiguë (à vérifier manuellement), ou si le fonds est fermé (sortie non disponible) |

Il n'y a pas de colonne « Titulaire » : quand la Consolidation subdivise les contrats par
titulaire (Monsieur / Madame / société...), chaque titulaire obtient son **propre bandeau**
(harmonisé en « M. Prénom NOM » / « Mme Prénom NOM » à partir du titre de Consolidation, qui
énonce explicitement le genre — jamais deviné), suivi de sa propre ligne d'en-tête répétée, puis
de ses catégories et ses fonds — pas une colonne au milieu d'un tableau commun. Un même fonds
détenu à la fois par Monsieur ET Madame apparaît donc dans les 2 sections, chacune avec sa propre
date d'investissement et son propre statut de pénalité, séparées par une ligne vide (non bordée)
pour qu'elles ne paraissent pas collées. Une seule ligne d'en-tête, sans bandeau, si le fichier ne
subdivise pas les contrats (client à titulaire unique). Si
aucun fonds détenu n'a de calendrier de rachat connu (ni fonds fermé), la feuille l'indique
clairement au lieu d'un tableau vide. Comme dans `ajout-calendrier-sortie.html`, la feuille
`Consolidation` reçoit aussi 2 colonnes (`Rachat — cash reçu`, `Pénalité de sortie`) collées juste
après sa dernière colonne réellement utile (une colonne « Mouvements en cours » existante, toujours
en dernière position, est repoussée de 2 colonnes vides après elles plutôt que d'être dépassée), en
formule vers « Calendrier de sortie » pour la même information dans les 2 pages. Le script Python
n'a pas d'équivalent au formulaire de saisie de dates dans le navigateur (voir plus haut) : les
dates s'y saisissent directement dans le fichier Excel généré, comme avant. La présentation
(bandeaux de catégorie beige, en-tête coloré, police) reprend celle de la feuille Consolidation —
voir la section dédiée à `ajout-calendrier-sortie.html` plus haut pour le détail.

### Générer / régénérer la feuille

```bash
python3 scripts/build_client_workbook.py [chemin_consolidation.xlsx] [chemin_sortie.xlsx]
```

Sans arguments, lit `source/ConsolidationTemplateAlthosAI_V5.xlsx` et écrit
`output/ConsolidationTemplateAlthosAI_V5_avec_calendrier.xlsx`. Pour l'utiliser sur le classeur
réel d'un client (avec les montants déjà saisis), passer son chemin en premier argument — les
données de la feuille `Consolidation` du client ne sont jamais modifiées, seules 2 colonnes sont
ajoutées à la suite (voir ci-dessus), en plus de la nouvelle feuille « Calendrier de sortie ».

⚠️ **Limite connue de ce script (Python/openpyxl uniquement, n'affecte pas `ajout-calendrier-sortie.html`)** :
openpyxl ne conserve pas les valeurs mises en cache des formules déjà présentes dans la feuille
`Consolidation` (ex. les totaux). Les formules elles-mêmes restent intactes et correctes ;
Excel les recalcule simplement à l'ouverture du fichier (comportement normal, pas une perte de
données) — mais le fichier peut afficher des cellules vides une fraction de seconde avant ce
recalcul automatique. Si ce détail gêne, préférez `ajout-calendrier-sortie.html`, qui conserve
les valeurs calculées intactes.

### Comment ça marche (toutes les colonnes calculées sont des formules Excel)

Le script ajoute aussi deux feuilles de données masquées, alimentées par
`Calendriers_de_fonds_Althos.xlsx` (même logique que `build_data.py`, voir plus haut) :

- **`BDD_Calendrier`** : une ligne par fonds / type (Souscription ou Rachat) / échéance mensuelle.
- **`BDD_Penalites`** : une ligne par fonds, avec la règle de pénalité réduite à au plus 4 paliers
  "seuil en mois → taux" + un taux "au-delà" (`build_tier_columns()`), pour rester calculable par
  une formule `IFS` en cascade sans jamais recourir à une formule matricielle (CSE). Pour les fonds
  fermés, 2 colonnes de texte libre séparées portent la durée de vie du fonds et sa durée de
  blocage/conservation (extraites verbatim du texte source, jamais reformulées).

Les 6 colonnes calculées visibles (les 5 dates de rachat + la pénalité) s'appuient sur des
colonnes de calcul intermédiaires masquées (prochaine échéance de rachat trouvée via
`COUNTIFS`/`MINIFS` à partir d'aujourd'hui, mois de détention, palier de pénalité applicable…).
Tout se recalcule automatiquement à l'ouverture du fichier, à la date du jour — rien n'est figé.

**⚠️ Limite connue de cette tâche : le recalcul automatique n'a pas pu être vérifié dans cet
environnement.** LibreOffice (utilisé habituellement pour valider les formules avant livraison)
n'a pas réussi à charger le moindre fichier `.xlsx` dans ce bac à sable — y compris les fichiers
sources originaux, jamais modifiés — ce qui pointe vers un problème d'infrastructure de la
session, indépendant de ce script. Les formules ont donc été relues manuellement avec soin
(cohérence des plages, parenthésage, gestion des cas vides/erreurs) plutôt que vérifiées par
exécution réelle. **À l'ouverture du fichier dans Excel ou LibreOffice Calc chez vous, vérifiez
qu'aucune cellule n'affiche `#NAME?`, `#REF!` ou `#VALUE!`** avant de vous fier au résultat ; le
cas échéant, signalez-le pour correction.
