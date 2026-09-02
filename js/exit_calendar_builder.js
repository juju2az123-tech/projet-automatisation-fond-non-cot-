/*
 * Ajoute, dans un classeur de consolidation client (feuille "Consolidation" : un fonds par
 * ligne, catégories en lignes surlignées, ISIN en colonne B, montants par contrat en colonnes
 * suivantes), une nouvelle feuille "Calendrier de sortie" — un tableau compact listant
 * UNIQUEMENT les fonds que le client détient réellement (montant total non nul) ET pour
 * lesquels un calendrier de RACHAT (sortie) officiel est connu dans la base Althos. Pas de
 * fonds coté, pas de fonds non détenu, pas de fonds sans calendrier de sortie : si un fonds n'a
 * rien à afficher, il n'apparaît simplement pas.
 *
 * Pour chaque fonds retenu, les valeurs de la prochaine échéance de rachat sont reprises
 * telles quelles depuis la base (une colonne par champ, pas de texte composé) :
 *   Titulaire | Fonds | ISIN | Date d'investissement | Rachat — ordre avant | Rachat — VL |
 *   Rachat — exécuté | Rachat — publié | Rachat — cash reçu | Pénalité de sortie
 *
 * La colonne Pénalité de sortie reste VIDE quand le client n'est plus concerné (délai de
 * pénalité dépassé) ; elle affiche un message dans tous les autres cas (aucune pénalité prévue,
 * pénalité en cours, information non renseignée, ou cas ambigu à vérifier manuellement).
 *
 * Présentation reprise de la feuille Consolidation elle-même (même police, mêmes couleurs) :
 * bandeaux de catégorie (beige, gras) au-dessus des fonds qu'ils regroupent, fonds sur fond
 * blanc, en-tête de tableau dans la même couleur que celui de Consolidation. La colonne
 * "Titulaire" indique via quel contrat (Monsieur / Madame / société / etc., lu sur la ligne
 * juste au-dessus de l'en-tête "Support" dans Consolidation) le fonds est détenu ; un même fonds
 * détenu par plusieurs titulaires donne une ligne par titulaire (dates et pénalité propres à
 * chacun). Reste vide si le fichier n'a qu'un seul titulaire (pas de subdivision par contrat).
 *
 * Port JavaScript (ExcelJS, exécuté dans le navigateur) de scripts/build_client_workbook.py —
 * même logique de calcul, mêmes formules. La structure du fichier source (numéros de ligne,
 * nombre de colonnes "Contrat") N'EST PAS supposée fixe : elle est détectée dynamiquement pour
 * s'adapter à n'importe quel classeur client réel.
 *
 * Dépend de window.FUNDS_DATA (data/funds_data.js) et de la bibliothèque ExcelJS globale
 * (vendor/exceljs.min.js), à charger avant ce fichier.
 */
(function (global) {
  "use strict";

  // -------------------------------------------------------------------------
  // Utilitaires
  // -------------------------------------------------------------------------

  function colLetter(n) {
    let s = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function toExcelDate(iso) {
    if (!iso) return null;
    const [y, m, d] = iso.split("-").map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  }

  function isSolidBold(cell) {
    const fill = cell.fill;
    const isSolid = !!(fill && fill.type === "pattern" && fill.pattern === "solid");
    const isBold = !!(cell.font && cell.font.bold);
    return isSolid && isBold;
  }

  /** "18/08/2026", au format numérique jour/mois/année. */
  function formatDateFr(d) {
    const p2 = (n) => String(n).padStart(2, "0");
    return `${p2(d.getDate())}/${p2(d.getMonth() + 1)}/${d.getFullYear()}`;
  }

  function cloneStyle(cell) {
    // Deep clone via JSON round-trip: ExcelJS style objects are plain data (no functions),
    // safe to clone this way and avoids two cells sharing (and mutating) the same object.
    return JSON.parse(JSON.stringify(cell.style || {}));
  }

  /** Couleur de quadrillage reprise telle quelle (référence de thème + teinte, PAS une valeur
   *  fixe) de la bordure déjà utilisée par l'en-tête de Consolidation — sa couleur réelle dépend
   *  du thème propre à chaque classeur (beige dans certains, bleu dans d'autres). Renvoie un
   *  bordure complète (4 côtés) prête à appliquer à une cellule. */
  function buildGridBorder(ws, headerRow) {
    const headerBorder = ws.getCell(headerRow, 1).border || {};
    const side = JSON.parse(JSON.stringify(
      headerBorder.left || headerBorder.top || headerBorder.right || headerBorder.bottom ||
      { style: "thin", color: { argb: "FFDDCCB8" } }
    ));
    side.style = "medium";
    return { top: side, left: side, bottom: side, right: side };
  }

  const PENALTY_PREFIXES = {
    ferme: "Fonds fermé : aucun rachat possible.",
    manuel: "À VÉRIFIER MANUELLEMENT : ",
    concerne1: "Pénalité de sortie : ",
    concerne2: "Durée de détention actuelle : XXX mois.",
  };

  /** Nombre de lignes à prévoir pour le texte de pénalité d'un fonds donné, une fois la mise en
   *  forme finale appliquée (2 phrases séparées par un retour à la ligne explicite pour "ferme"
   *  et les pénalités actives datées) — calculé à partir de la longueur RÉELLE du texte de la
   *  base (connue à la génération), pas devinée : un texte source inhabituellement long (ça
   *  arrive) donne une ligne plus haute plutôt qu'un texte coupé. */
  function estimatePenaltyLines(kind, rawLen, dureeLen, charsPerLine) {
    const linesFor = (len) => Math.max(1, Math.ceil(len / charsPerLine));
    if (kind === "ferme") {
      let lines = linesFor(PENALTY_PREFIXES.ferme.length);
      if (dureeLen) lines += linesFor(dureeLen);
      return lines;
    }
    if (kind === "manuel") {
      return linesFor(PENALTY_PREFIXES.manuel.length + rawLen);
    }
    if (kind === "seuil" || kind === "soft" || kind === "degressif") {
      return linesFor(PENALTY_PREFIXES.concerne1.length + rawLen) + linesFor(PENALTY_PREFIXES.concerne2.length);
    }
    return 1; // aucune / inconnue / date manquante / date future : message court
  }

  /** Hauteur de ligne (en points) pour `lines` lignes de texte à la taille de police `fontSize`
   *  — approximation généreuse du ratio hauteur de ligne / taille de police (~1.5, contre ~1.2
   *  pour une police système classique) : la police réelle utilisée dans ces classeurs
   *  (Montserrat) est plus large et rend visuellement plus haute qu'une police par défaut, donc
   *  la marge est volontairement large pour ne jamais couper le texte de justesse. */
  function heightForLines(lines, fontSize) {
    return Math.round(lines * (fontSize * 1.6 + 4));
  }

  // Volontairement plus haut qu'un minimum "juste" : sans moteur de rendu disponible pour
  // vérifier pixel par pixel, on préfère une ligne visiblement plus haute que nécessaire à un
  // texte coupé dans le classeur final.
  const DEFAULT_PENALTY_ROW_HEIGHT = 75;

  // -------------------------------------------------------------------------
  // Pénalité -> 9 valeurs numériques pour une cascade IFS (identique à
  // build_tier_columns() dans scripts/build_client_workbook.py)
  // -------------------------------------------------------------------------

  function buildTierColumns(pen) {
    const kind = pen.kind;
    let display;
    if (kind === "seuil" || kind === "soft") {
      const t = pen.tiers[0];
      display = [
        { max: t.maxMonths, rate: t.rate },
        { max: null, rate: 0 },
      ];
    } else if (kind === "degressif") {
      display = pen.tiers.map((t) => ({ max: t.maxMonths, rate: t.rate }));
    } else {
      return [0, 0, 0, 0, 0, 0, 0, 0, 0];
    }
    let conditions = display.slice(0, -1);
    const tailRate = display[display.length - 1].rate;
    while (conditions.length < 4) {
      const lastMax = conditions.length ? conditions[conditions.length - 1].max : 0;
      conditions.push({ max: lastMax, rate: tailRate });
    }
    conditions = conditions.slice(0, 4);
    const out = [];
    conditions.forEach((c) => {
      out.push(c.max);
      out.push(c.rate);
    });
    out.push(tailRate);
    return out; // [Max1,Rate1,Max2,Rate2,Max3,Rate3,Max4,Rate4,Rate5]
  }

  // -------------------------------------------------------------------------
  // Feuilles de données (BDD_Calendrier / BDD_Penalites)
  // -------------------------------------------------------------------------

  function writeBddCalendrier(workbook, calendar) {
    const ws = workbook.addWorksheet("BDD_Calendrier");
    const headers = ["ISIN", "Nom", "Type", "Cutoff", "Valorisation", "Execution",
      "PublicationVL", "ReglementCash", "CleComposite"];
    ws.addRow(headers);

    Object.keys(calendar).forEach((isin) => {
      const byType = calendar[isin];
      Object.keys(byType).forEach((type_) => {
        const entries = byType[type_];
        const expandedTypes = type_ === "Souscription et rachat" ? ["Souscription", "Rachat"] : [type_];
        expandedTypes.forEach((etype) => {
          entries.forEach((e) => {
            const cutoff = e.cutoff;
            const key = cutoff ? `${isin}|${etype}|${cutoff}` : "";
            const row = ws.addRow([
              isin, "", etype,
              toExcelDate(e.cutoff), toExcelDate(e.valorisation), toExcelDate(e.execution),
              toExcelDate(e.publicationVL), toExcelDate(e.reglementCash), key,
            ]);
            [4, 5, 6, 7, 8].forEach((c) => { row.getCell(c).numFmt = "dd/mm/yyyy"; });
          });
        });
      });
    });

    const widths = [14, 4, 12, 12, 12, 12, 12, 12, 26];
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.getRow(1).font = { bold: true };
    ws.state = "hidden";
    return ws;
  }

  function writeBddPenalites(workbook, funds) {
    const ws = workbook.addWorksheet("BDD_Penalites");
    const headers = ["ISIN", "Nom", "Kind", "Max1", "Rate1", "Max2", "Rate2",
      "Max3", "Rate3", "Max4", "Rate4", "Rate5", "RawText", "DureeVie"];
    ws.addRow(headers);

    funds.forEach((f) => {
      if (!f.isin) return;
      const pen = f.penalite || { kind: "inconnue", raw: null, tiers: [] };
      const tiers9 = buildTierColumns(pen);
      ws.addRow([f.isin, f.nom, pen.kind, ...tiers9, pen.raw || "", pen.dureeVie || ""]);
    });

    const widths = [14, 34, 11, 7, 7, 7, 7, 7, 7, 7, 7, 7, 60, 60];
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.getRow(1).font = { bold: true };
    ws.state = "hidden";
    return ws;
  }

  // -------------------------------------------------------------------------
  // Détection dynamique de la structure de la feuille "Consolidation"
  // -------------------------------------------------------------------------

  function findConsolidationSheet(workbook) {
    let ws = workbook.getWorksheet("Consolidation");
    if (ws) return ws;
    // Repli : première feuille dont la cellule A d'une des 15 premières lignes vaut "Support"
    for (const candidate of workbook.worksheets) {
      for (let r = 1; r <= 15; r++) {
        const v = (candidate.getCell(r, 1).value || "").toString().trim();
        if (v === "Support") return candidate;
      }
    }
    return null;
  }

  function findHeaderRow(ws) {
    for (let r = 1; r <= 15; r++) {
      const v = (ws.getCell(r, 1).value || "").toString().trim();
      if (v === "Support") return r;
    }
    return null;
  }

  /** Colonne dont l'en-tête (sur headerRow) vaut "TOTAL" (montant total détenu par fonds). */
  function findTotalColumn(ws, headerRow) {
    const lastCol = ws.actualColumnCount || ws.columnCount;
    for (let c = 3; c <= lastCol; c++) {
      const v = (ws.getCell(headerRow, c).value || "").toString().trim().toUpperCase();
      if (v === "TOTAL") return c;
    }
    return null;
  }

  /**
   * Classe chaque ligne : 'category' (bandeau, à ignorer), 'blank', ou 'fund'. Associe aussi à
   * chaque ligne fonds le libellé de la dernière catégorie (bandeau) rencontrée au-dessus d'elle,
   * pour pouvoir reproduire les mêmes bandeaux de catégorie dans la feuille générée.
   */
  /** Vrai si `v` (valeur d'une cellule TOTAL) correspond à une vraie agrégation de fonds —
   *  =SUM(...) sur ses colonnes de contrat, ou un nombre déjà calculé — plutôt qu'une simple
   *  formule de renvoi vers une autre cellule (=A8, =IFERROR(U312,"/")...), utilisée par certains
   *  petits tableaux récapitulatifs hors-tableau plus bas sur la feuille ("dont actions
   *  européennes : 15 %"...) qui ne sont PAS des lignes de fonds malgré un texte en colonne A. */
  function isSumTotal(v) {
    if (typeof v === "number") return true;
    if (typeof v === "string") return /^=?SUM\(/i.test(v.trim());
    if (v && typeof v === "object") {
      if (typeof v.formula === "string") return /^SUM\(/i.test(v.formula.trim());
      // Cellule "esclave" d'une formule partagée (Excel factorise le texte d'une formule
      // identique recopiée sur toute une colonne — ExcelJS ne porte alors le texte que sur la
      // cellule maîtresse). Comme la colonne TOTAL n'est JAMAIS que des =SUM(...) recopiées
      // (jamais un renvoi de cellule ponctuel comme "=A8", qui ne serait de toute façon jamais
      // reconnu "partageable" par Excel avec une plage =SUM contiguë), une cellule esclave ici
      // est nécessairement, elle aussi, une vraie agrégation.
      if (typeof v.sharedFormula === "string") return true;
    }
    return false;
  }

  function classifyRows(ws, headerRow, totalCol) {
    const categoryRows = [];
    const fundRows = [];
    const rowToCategory = {};
    let currentCategory = null;
    const lastRow = ws.actualRowCount || ws.rowCount;
    // headerRow+1 = ligne de total général (pas une catégorie, pas un fonds) -> ignorée
    for (let r = headerRow + 2; r <= lastRow; r++) {
      const a = ws.getCell(r, 1);
      if (a.value === null || a.value === undefined || a.value === "") continue;
      if (isSolidBold(a)) {
        categoryRows.push(r);
        currentCategory = String(a.value).trim();
      } else {
        if (totalCol && !isSumTotal(ws.getCell(r, totalCol).value)) continue;
        fundRows.push(r);
        rowToCategory[r] = currentCategory;
      }
    }
    return { categoryRows, fundRows, rowToCategory };
  }

  /** Dernière colonne portant un intitulé sur la ligne d'en-tête (pour savoir où ajouter des
   *  colonnes à la suite, sans écraser une colonne existante — ex. "Risque / 7"). Balaie une
   *  large plage fixe plutôt que `actualColumnCount` : ce dernier peut sous-compter quand la
   *  dernière colonne utile est une cellule "esclave" d'une fusion (ExcelJS lui fait quand même
   *  refléter la valeur de la cellule maîtresse en lecture — voir isMergedFollower ci-dessous). */
  function findLastHeaderColumn(ws, headerRow) {
    const upper = Math.max(ws.columnCount || 0, ws.actualColumnCount || 0, 200);
    let last = 1;
    for (let c = 1; c <= upper; c++) {
      const v = ws.getCell(headerRow, c).value;
      if (v !== null && v !== undefined && String(v).trim() !== "") last = c;
    }
    return last;
  }

  /** Vrai si la cellule est "esclave" d'une fusion (pas la cellule maîtresse en haut à gauche) :
   *  y écrire une valeur la redirigerait silencieusement vers la cellule maîtresse dans ExcelJS
   *  (contrairement à openpyxl, qui refuse carrément l'écriture sur ce type de cellule). */
  function isMergedFollower(ws, row, col) {
    const cell = ws.getCell(row, col);
    return !!(cell.isMerged && cell.master && cell.master.address !== cell.address);
  }

  /** Première colonne à partir de `from` qui n'est pas une cellule esclave d'une fusion sur
   *  cette ligne — pour ne jamais faire atterrir une nouvelle colonne au milieu d'une fusion
   *  existante (ex. "Mouvements en cours" fusionnée sur 2 colonnes dans certains fichiers). */
  function nextSafeColumn(ws, row, from) {
    let c = from;
    while (isMergedFollower(ws, row, c)) c++;
    return c;
  }

  /** Montant total détenu sur une ligne fonds : somme des colonnes "Contrat" (3..totalCol-1). */
  function holdingAmount(ws, row, totalCol) {
    if (!totalCol) return null; // inconnu : on ne filtre pas sur le montant
    let sum = 0;
    for (let c = 3; c < totalCol; c++) {
      const v = ws.getCell(row, c).value;
      if (typeof v === "number") sum += v;
    }
    return sum;
  }

  /**
   * Libellé du titulaire (Monsieur / Madame / société / prénom...) associé à chaque colonne
   * "Contrat", lu sur la ligne juste au-dessus de l'en-tête (souvent fusionnée sur plusieurs
   * colonnes contrat). Absent ou vide pour un client à titulaire unique (pas de subdivision) :
   * dans ce cas toutes les colonnes retombent sur "" et les montants sont simplement regroupés.
   */
  function detectOwnerLabels(ws, headerRow, totalCol) {
    const labels = {};
    const ownerRow = headerRow - 1;
    if (!totalCol || ownerRow < 1) return labels;
    for (let c = 3; c < totalCol; c++) {
      const v = ws.getCell(ownerRow, c).value;
      labels[c] = v === null || v === undefined ? "" : String(v).trim();
    }
    return labels;
  }

  /** Extrait {"nom complet (dans l'ordre du titre)": "M."|"Mme"} depuis le texte du titre de
   *  Consolidation (ex. "Consolidation Monsieur Gérard LENO**** et Madame Marie MAUN***"), pour
   *  harmoniser les bandeaux de titulaire de "Calendrier de sortie" uniquement (jamais
   *  Consolidation elle-même, qui garde son propre texte tel quel). Ne devine jamais un genre :
   *  ne renvoie que ce que le titre énonce explicitement. */
  function extractOwnerCivilities(titleText) {
    const out = {};
    if (!titleText) return out;
    const re = /(Monsieur|Madame)\s+([^,;]+?)(?=\s+(?:et|Monsieur|Madame)\b|$)/gi;
    let m;
    while ((m = re.exec(String(titleText))) !== null) {
      const civ = /^monsieur$/i.test(m[1]) ? "M." : "Mme";
      out[m[2].trim()] = civ;
    }
    return out;
  }

  /** Reformate un libellé de titulaire "NOM Prénom" (tel que lu dans Consolidation) en
   *  "M. Prénom NOM" / "Mme Prénom NOM" quand le titre de Consolidation permet une
   *  correspondance sans ambiguïté (mêmes mots, sans deviner un genre) ; renvoie le libellé
   *  d'origine inchangé sinon (ex. une société). */
  function harmonizeOwnerLabel(label, civilities) {
    if (!label) return label;
    const labelTokens = label.split(/\s+/).filter(Boolean).map((t) => t.toLowerCase()).sort().join("|");
    for (const name of Object.keys(civilities)) {
      const nameTokens = name.split(/\s+/).filter(Boolean).map((t) => t.toLowerCase()).sort().join("|");
      if (labelTokens && labelTokens === nameTokens) return `${civilities[name]} ${name}`;
    }
    return label;
  }

  /** Répartit le montant détenu d'une ligne fonds par titulaire (colonne "" si non subdivisé). */
  function holdingByOwner(ws, row, totalCol, ownerLabels) {
    if (!totalCol) return [["", null]]; // structure inconnue : une seule ligne, montant indéterminé
    const amounts = new Map();
    for (let c = 3; c < totalCol; c++) {
      const v = ws.getCell(row, c).value;
      if (typeof v !== "number") continue;
      const label = ownerLabels[c] || "";
      amounts.set(label, (amounts.get(label) || 0) + v);
    }
    return [...amounts.entries()];
  }

  // -------------------------------------------------------------------------
  // Construction de la feuille "Calendrier de sortie"
  // -------------------------------------------------------------------------

  const HELPER_NAMES = [
    "has_sortie", "next_cutoff_sortie", "next_val_sortie", "next_exec_sortie",
    "next_pub_sortie", "next_cash_sortie",
    "months_held", "pen_found", "kind", "raw", "duree_vie",
    "max1", "rate1", "max2", "rate2", "max3", "rate3", "max4", "rate4", "rate5",
    "rate_now",
  ];
  const HELPER_FIRST_COL = 11; // K (colonnes visibles jusqu'en I : Fonds, ISIN, Date, 5 dates de rachat, Pénalité)

  function helperCol(name) {
    return colLetter(HELPER_FIRST_COL + HELPER_NAMES.indexOf(name));
  }

  /** Un fonds n'a de calendrier "de sortie" que s'il a des échéances de type Rachat. */
  function hasRachatCalendar(calendar, isin) {
    const byType = calendar[isin];
    return !!(byType && (byType["Rachat"] || byType["Souscription et rachat"]));
  }

  function buildExitSheet(workbook, srcWs, headerRow, calendar, calLastRow, penLastRow, fundsByIsin) {
    const cal = (col) => `BDD_Calendrier!$${col}$2:$${col}$${calLastRow}`;
    const pen = (col) => `BDD_Penalites!$${col}$2:$${col}$${penLastRow}`;

    const ws = workbook.addWorksheet("Calendrier de sortie");
    // Quadrillage Excel par défaut désactivé, comme sur Consolidation : sans ça, le fond gris
    // clair du quadrillage standard reste visible autour des bordures de couleur ajoutées plus
    // bas, et le rendu ne paraît pas "blanc propre" comme sur Consolidation. Vue "aperçu des
    // sauts de page" (style: "pageBreakPreview", propriété interne d'ExcelJS pour l'attribut XML
    // <sheetView view="pageBreakPreview">), comme sur Consolidation : c'est ce qui affiche
    // automatiquement le contour bleu de la zone d'impression en vue Normale.
    ws.views = [{ showGridLines: false, style: "pageBreakPreview" }];

    // 1) Repérage des fonds à retenir : détenus (montant non nul, réparti par titulaire de
    //    contrat) ET dotés d'un calendrier de RACHAT connu dans la base Althos. Tout le reste
    //    (fonds cotés, fonds non cotés non détenus, fonds sans calendrier de sortie) n'apparaît
    //    pas dans la feuille. Un même fonds détenu par plusieurs titulaires (ex. Monsieur ET
    //    Madame, chacun via son propre contrat) donne une ligne par titulaire, pour permettre une
    //    date d'investissement et un statut de pénalité propres à chacun.
    const totalCol = findTotalColumn(srcWs, headerRow);
    const { fundRows, rowToCategory } = classifyRows(srcWs, headerRow, totalCol);
    const ownerLabels = detectOwnerLabels(srcWs, headerRow, totalCol);

    const selected = [];
    fundRows.forEach((srcRow) => {
      const isinRaw = srcWs.getCell(srcRow, 2).value;
      const isin = typeof isinRaw === "string" ? isinRaw.trim() : isinRaw;
      const fund = isin ? fundsByIsin.get(isin) : null;
      // Un fonds fermé (aucune sortie possible hors cas exceptionnel) est toujours retenu même
      // sans calendrier de rachat : c'est justement l'information à signaler au conseiller.
      const isFerme = fund && fund.penalite && fund.penalite.kind === "ferme";
      if (!fund || !(isFerme || hasRachatCalendar(calendar, isin))) return;
      const nom = srcWs.getCell(srcRow, 1).value || fund.nom;
      const category = rowToCategory[srcRow] || "";
      // Seuls les fonds à règle de pénalité datée (seuil/soft/dégressif) ont besoin d'une date
      // d'investissement pour statuer : "aucune"/"inconnue"/"manuel"/"ferme" ne dépendent pas de
      // la date. C'est cette liste-là qu'on demandera au conseiller de renseigner.
      const penKind = fund.penalite && fund.penalite.kind;
      const needsDate = penKind === "seuil" || penKind === "soft" || penKind === "degressif";
      holdingByOwner(srcWs, srcRow, totalCol, ownerLabels).forEach(([owner, amount]) => {
        const isHeld = amount === null || Math.abs(amount) > 0.005;
        if (!isHeld) return;
        selected.push({ isin, nom, category, owner, amount, needsDate, penalite: fund.penalite });
      });
    });

    // Regroupement par titulaire : un tableau complètement séparé par titulaire (Monsieur /
    // Madame / société...), plutôt qu'une colonne "Titulaire" au milieu d'un tableau commun —
    // pour que le conseiller voie d'un coup d'œil tout ce qui est détenu par chacun. L'ordre de
    // ces titulaires suit l'ordre de leurs colonnes "Contrat" dans Consolidation (premier titulaire
    // rencontré = premier affiché). Si le fichier n'a qu'un seul titulaire (pas de subdivision par
    // contrat, ownerLabels toujours ""), on garde un unique tableau, sans bandeau de titulaire.
    const ownerOrder = [];
    const byOwner = new Map();
    selected.forEach((item) => {
      const key = item.owner || "";
      if (!byOwner.has(key)) { byOwner.set(key, []); ownerOrder.push(key); }
      byOwner.get(key).push(item);
    });
    const showOwnerHeadings = ownerOrder.length > 1 || (ownerOrder.length === 1 && ownerOrder[0] !== "");
    // Harmonisation "M. Prénom NOM" / "Mme Prénom NOM" (uniquement sur cette feuille, jamais sur
    // Consolidation elle-même) à partir du titre de Consolidation, qui énonce explicitement
    // "Monsieur"/"Madame" pour chaque titulaire — jamais deviné. Le bandeau de titulaire reprend
    // le même style que l'en-tête du tableau (pas le style, souvent plus petit, de la ligne où
    // Consolidation affiche "Monsieur"/"Madame" au-dessus de ses colonnes contrat) : c'est un
    // bandeau pleine largeur, pas une étiquette de sous-colonne.
    const ownerCivilities = showOwnerHeadings ? extractOwnerCivilities(srcWs.getCell(1, 1).value) : {};

    // 2) Présentation "à la Althos" : bandeau de titre + sous-titre daté (repris tel quel des
    //    lignes 1 et 3 de Consolidation — même couleur, même police, même mise en italique),
    //    en-tête de tableau dans le même bleu que Consolidation et le titre, bandeaux de
    //    catégorie en beige (repris de la couleur des catégories dans Consolidation), fonds sur
    //    fond blanc, quadrillage fin en beige. Police commune à toute la feuille (celle de
    //    l'en-tête Consolidation, pas un choix arbitraire).
    const headerStyle = cloneStyle(srcWs.getCell(headerRow, 1));
    const baseFontName = (headerStyle.font && headerStyle.font.name) || "Calibri";
    const baseFontSize = (headerStyle.font && headerStyle.font.size) || 10;
    const dataFont = { name: baseFontName, size: baseFontSize, bold: false };
    const boldDataFont = { name: baseFontName, size: baseFontSize, bold: true };

    const titleStyle = cloneStyle(srcWs.getCell(1, 1));
    const subtitleStyle = cloneStyle(srcWs.getCell(3, 1));
    // Style de bandeau de catégorie : repris tel quel de la première ligne de catégorie trouvée
    // dans la Consolidation (même couleur beige, même police en gras).
    const { categoryRows } = classifyRows(srcWs, headerRow, totalCol);
    const categoryStyle = categoryRows.length ? cloneStyle(srcWs.getCell(categoryRows[0], 1)) : null;
    const gridBorder = buildGridBorder(srcWs, headerRow);

    const headers = [
      "Fonds", "ISIN", "Date d'investissement",
      "Rachat — ordre avant", "Rachat — VL", "Rachat — exécuté",
      "Rachat — publié", "Rachat — cash reçu", "Pénalité de sortie",
    ];
    // 1: titre, 2: (vide), 3: sous-titre daté, 4: (vide), 5: début du tableau — soit directement
    // la ligne d'en-tête (un seul titulaire), soit le 1er bandeau de titulaire suivi de sa propre
    // ligne d'en-tête (plusieurs titulaires).
    const HEADER_ROW_OUT = 5;
    const FIRST_DATA_ROW = HEADER_ROW_OUT + 1;

    const titleCell = ws.getCell(1, 1);
    titleCell.value = "Calendrier des délais de sortie";
    titleCell.style = JSON.parse(JSON.stringify(titleStyle));
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.mergeCells(1, 1, 1, headers.length);
    ws.getRow(1).height = srcWs.getRow(1).height || 22.5;

    const subtitleCell = ws.getCell(3, 1);
    subtitleCell.value = formatDateFr(new Date());
    subtitleCell.style = JSON.parse(JSON.stringify(subtitleStyle));
    subtitleCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.mergeCells(3, 1, 3, headers.length);
    ws.getRow(3).height = srcWs.getRow(3).height || 14;

    // Ligne d'en-tête ("Fonds | ISIN | ..."), réutilisable : écrite une seule fois si le fichier
    // n'a qu'un seul titulaire, ou répétée juste sous chaque bandeau de titulaire sinon — pour
    // que chaque section du tableau soit lisible seule, sans remonter chercher les intitulés.
    function writeHeaderRow(atRow) {
      headers.forEach((label, i) => {
        const cell = ws.getCell(atRow, i + 1);
        cell.value = label;
        cell.style = JSON.parse(JSON.stringify(headerStyle));
        cell.alignment = { horizontal: "center", vertical: "middle", wrapText: false };
        cell.border = gridBorder;
      });
      ws.getRow(atRow).height = 22;
    }
    if (!showOwnerHeadings) writeHeaderRow(HEADER_ROW_OUT);
    const PENALTY_COL_WIDTH = 63;
    // Facteur 0.7 : la police réelle (Montserrat) est plus large qu'une police système classique
    // et le retour à la ligne se fait sur des mots entiers (jamais exactement à la largeur max),
    // donc le nombre de caractères qui tiennent réellement sur une ligne est nettement inférieur
    // à la largeur de colonne brute — mieux vaut sur-estimer le nombre de lignes que couper le texte.
    const CHARS_PER_LINE = Math.round(PENALTY_COL_WIDTH * 0.55);
    const widths = [37, 16, 22, 22, 15, 19, 18, 20, PENALTY_COL_WIDTH];
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    HELPER_NAMES.forEach((name) => { ws.getColumn(colNumOf(helperCol(name))).hidden = true; });

    if (!selected.length) {
      ws.getCell(FIRST_DATA_ROW, 1).value =
        "Aucun fonds avec calendrier de sortie (rachat) connu n'est actuellement détenu par ce " +
        "client (montant total nul, ou fonds hors périmètre \"fonds non cotés suivis\").";
      ws.mergeCells(FIRST_DATA_ROW, 1, FIRST_DATA_ROW, headers.length);
      ws.getCell(FIRST_DATA_ROW, 1).alignment = { wrapText: true, vertical: "middle" };
      ws.getRow(FIRST_DATA_ROW).height = 30;
      applyPrintSetup(ws, FIRST_DATA_ROW, headers.length);
      return { ws, selectedCount: 0, fundRowsScanned: fundRows.length, rowsNeedingDate: [], firstDataRow: null, lastDataRow: null };
    }

    // 3) Un bandeau par titulaire suivi de sa propre ligne d'en-tête (bandeaux de catégorie beige,
    //    puis une ligne par fonds retenu, avec les formules de calcul) — chaque titulaire séparé
    //    du suivant par 1 ligne vide, pour que 2 tableaux distincts ne paraissent pas collés.
    let r = showOwnerHeadings ? HEADER_ROW_OUT : FIRST_DATA_ROW;
    const rowsNeedingDate = [];
    ownerOrder.forEach((ownerKey, ownerIdx) => {
      if (showOwnerHeadings) {
        const heading = harmonizeOwnerLabel(ownerKey, ownerCivilities);
        ws.getCell(r, 1).value = heading || "Autres titulaires";
        for (let c = 1; c <= headers.length; c++) {
          ws.getCell(r, c).style = JSON.parse(JSON.stringify(headerStyle));
        }
        ws.mergeCells(r, 1, r, headers.length);
        ws.getCell(r, 1).alignment = { horizontal: "center", vertical: "middle" };
        ws.getRow(r).height = 20;
        r += 1;
        writeHeaderRow(r);
        r += 1;
      }

      let currentCategory; // undefined != "" : force le 1er bandeau même si catégorie ""
      byOwner.get(ownerKey).forEach((item) => {
        if (categoryStyle && item.category !== currentCategory) {
          currentCategory = item.category;
          ws.getCell(r, 1).value = currentCategory || "Autres fonds";
          for (let c = 1; c <= headers.length; c++) {
            ws.getCell(r, c).style = JSON.parse(JSON.stringify(categoryStyle));
            ws.getCell(r, c).border = gridBorder;
          }
          ws.mergeCells(r, 1, r, headers.length);
          ws.getCell(r, 1).alignment = { vertical: "middle" };
          ws.getRow(r).height = 20;
          r += 1;
        }

        for (let c = 1; c <= headers.length; c++) {
          ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
          ws.getCell(r, c).border = gridBorder;
        }
        ws.getCell(r, 1).value = item.nom;
        ws.getCell(r, 2).value = item.isin;
        [1, 2].forEach((c) => { ws.getCell(r, c).font = dataFont; ws.getCell(r, c).alignment = { vertical: "middle" }; });

        const dateCell = ws.getCell(r, 3);
        dateCell.numFmt = "dd/mm/yyyy";
        dateCell.font = dataFont;
        dateCell.alignment = { horizontal: "center", vertical: "middle" };

        if (item.needsDate) {
          rowsNeedingDate.push({ row: r, titulaire: item.owner || "", fonds: item.nom, isin: item.isin, categorie: currentCategory || "" });
        }

        const b = `"${item.isin}"`;
        const c = `$C${r}`;

        const L = helperCol("has_sortie"), M = helperCol("next_cutoff_sortie"), N = helperCol("next_val_sortie"),
          Nx = helperCol("next_exec_sortie"), Np = helperCol("next_pub_sortie"), O = helperCol("next_cash_sortie");
        const Q = helperCol("months_held");
        const R = helperCol("pen_found");
        const S = helperCol("kind");
        const Tc = helperCol("raw");
        const Dv = helperCol("duree_vie");
        const U1 = helperCol("max1"), V1 = helperCol("rate1"), W1 = helperCol("max2"), X1 = helperCol("rate2"),
          Y1 = helperCol("max3"), Z1 = helperCol("rate3"), AA1 = helperCol("max4"), AB1 = helperCol("rate4"), AC1 = helperCol("rate5");
        const AD1 = helperCol("rate_now");

        setF(ws, `${L}${r}`, `COUNTIFS(${cal("A")},${b},${cal("C")},"Rachat")`);
        setF(ws, `${M}${r}`, `IF(${L}${r}=0,"",IFERROR(_xlfn.MINIFS(${cal("D")},${cal("A")},${b},${cal("C")},"Rachat",${cal("D")},">="&TODAY()),""))`);
        // VL / exécuté / publié / cash reçu de CETTE échéance précise : on réutilise MINIFS avec
        // une égalité exacte sur la date de cut-off déjà trouvée (au lieu d'une reconstruction de
        // clé texte + MATCH, plus fragile) — même mécanisme que ${M} ci-dessus, qui fonctionne de
        // façon fiable. MINIFS ignore les cellules vides : si ce champ précis n'est pas renseigné
        // dans la base pour cette échéance, MINIFS ne trouve aucune valeur numérique et renvoie 0
        // (jamais une vraie erreur) — sans la vérification "=0" ci-dessous, Excel afficherait ce 0
        // comme une date, "00/01/1900", au lieu de laisser la cellule vide.
        const minifsField = (col) =>
          `IF(${M}${r}="","",IFERROR(IF(_xlfn.MINIFS(${cal(col)},${cal("A")},${b},${cal("C")},"Rachat",${cal("D")},${M}${r})=0,"",` +
          `_xlfn.MINIFS(${cal(col)},${cal("A")},${b},${cal("C")},"Rachat",${cal("D")},${M}${r})),""))`;
        setF(ws, `${N}${r}`, minifsField("E"));
        setF(ws, `${Nx}${r}`, minifsField("F"));
        setF(ws, `${Np}${r}`, minifsField("G"));
        setF(ws, `${O}${r}`, minifsField("H"));

        setF(ws, `${Q}${r}`, `IF(${c}="","",IF(${c}>TODAY(),"FUTUR",DATEDIF(${c},TODAY(),"m")))`);
        setF(ws, `${R}${r}`, `COUNTIF(${pen("A")},${b})`);
        setF(ws, `${S}${r}`, `IF(${R}${r}=0,"inconnue",INDEX(${pen("C")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${Tc}${r}`, `IF(${R}${r}=0,"",INDEX(${pen("M")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${Dv}${r}`, `IF(${R}${r}=0,"",INDEX(${pen("N")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${U1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("D")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${V1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("E")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${W1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("F")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${X1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("G")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${Y1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("H")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${Z1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("I")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${AA1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("J")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${AB1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("K")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${AC1}${r}`, `IF(${R}${r}=0,0,INDEX(${pen("L")},MATCH(${b},${pen("A")},0)))`);
        setF(ws, `${AD1}${r}`, `IF(OR(${c}="",${Q}${r}="FUTUR"),"",_xlfn.IFS(${Q}${r}<${U1}${r},${V1}${r},${Q}${r}<${W1}${r},${X1}${r},${Q}${r}<${Y1}${r},${Z1}${r},${Q}${r}<${AA1}${r},${AB1}${r},TRUE(),${AC1}${r}))`);

        // Colonnes visibles D..H : valeurs reprises telles quelles de la base (une par champ).
        [["D", M], ["E", N], ["F", Nx], ["G", Np], ["H", O]].forEach(([col, helper]) => {
          setF(ws, `${col}${r}`, `${helper}${r}`);
          const cell = ws.getCell(`${col}${r}`);
          cell.numFmt = "dd/mm/yyyy";
          cell.font = dataFont;
          cell.alignment = { horizontal: "center", vertical: "middle" };
        });

        // Pénalité de sortie : vide dès qu'il n'y a rien d'actionnable à signaler — pas de
        // pénalité prévue pour ce fonds (kind="aucune"), aucune pénalité renseignée dans la base
        // (kind="inconnue", traité comme "pas de pénalité"), ou délai de pénalité dépassé. Un
        // message n'apparaît que dans les cas où le conseiller doit agir : fonds fermé (sortie
        // impossible hors cas exceptionnel), règle ambiguë à vérifier à la main, ou pénalité
        // active à signaler au client.
        setF(ws, `I${r}`,
          `_xlfn.IFS(` +
          `${S}${r}="ferme","Fonds fermé : aucun rachat possible."&IF(${Dv}${r}<>"",CHAR(10)&${Dv}${r},""),` +
          `${S}${r}="manuel","À VÉRIFIER MANUELLEMENT : "&${Tc}${r},` +
          `${S}${r}="aucune","",` +
          `${S}${r}="inconnue","",` +
          `${c}="","Saisir une date d'investissement pour statuer sur la pénalité.",` +
          `${Q}${r}="FUTUR","Date d'investissement postérieure à aujourd'hui — vérifier la saisie.",` +
          `${AD1}${r}>0,"Pénalité de sortie : "&${Tc}${r}&CHAR(10)&"Durée de détention actuelle : "&${Q}${r}&" mois.",` +
          `TRUE(),"")`
        );
        const penCell = ws.getCell(`I${r}`);
        // Une formule Excel ne peut jamais renvoyer un texte enrichi (gras partiel) : tout ce
        // qu'affiche cette cellule (message de pénalité ou de fermeture) est mis en gras en bloc.
        penCell.font = boldDataFont;
        penCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
        // Hauteur de 60pt par défaut (confortable pour l'immense majorité des messages, 2-3
        // lignes) ; augmentée seulement si le texte réel de la base pour ce fonds précis est
        // inhabituellement long (ça arrive), pour ne jamais le couper.
        const penInfo = item.penalite || { kind: "inconnue", raw: null };
        const lines = estimatePenaltyLines(penInfo.kind, (penInfo.raw || "").length, (penInfo.dureeVie || "").length, CHARS_PER_LINE);
        ws.getRow(r).height = Math.max(DEFAULT_PENALTY_ROW_HEIGHT, heightForLines(lines, baseFontSize));
        r += 1;
      });

      // 1 ligne vide (non bordée) entre 2 titulaires, pour que 2 tableaux distincts ne
      // paraissent pas collés l'un à l'autre.
      if (showOwnerHeadings && ownerIdx < ownerOrder.length - 1) r += 1;
    });

    applyPrintSetup(ws, r - 1, headers.length);
    return {
      ws, selectedCount: selected.length, fundRowsScanned: fundRows.length, rowsNeedingDate,
      firstDataRow: FIRST_DATA_ROW, lastDataRow: r - 1,
    };
  }

  /**
   * Ajoute 2 colonnes à la suite du tableau Consolidation lui-même — "Rachat — cash reçu" et
   * "Pénalité de sortie" — pour que le conseiller ait l'essentiel sous les yeux sans changer
   * d'onglet. Ce sont des formules qui vont chercher, par ISIN, la valeur déjà calculée dans la
   * feuille "Calendrier de sortie" : une seule source de vérité, jamais recalculée en double. Ne
   * remplit que les fonds réellement détenus (même filtre que "Calendrier de sortie"). Si un même
   * fonds est détenu via plusieurs titulaires (plusieurs lignes dans "Calendrier de sortie"), la
   * ligne unique de Consolidation ne peut représenter qu'un seul statut : la première trouvée par
   * MATCH — la répartition détaillée par titulaire reste disponible dans "Calendrier de sortie".
   */
  /** Colonne du premier en-tête dont le texte correspond exactement (après trim) à `text`, ou
   *  null si absent — pour repérer une colonne "connue" comme "Mouvements en cours" par son
   *  intitulé plutôt que par une position codée en dur. */
  function findColumnByHeaderText(ws, headerRow, text) {
    const upper = Math.max(ws.columnCount || 0, ws.actualColumnCount || 0, 200);
    for (let c = 1; c <= upper; c++) {
      const v = ws.getCell(headerRow, c).value;
      if (v !== null && v !== undefined && String(v).trim() === text) return c;
    }
    return null;
  }

  /** Dernière colonne du bloc fusionné (horizontalement) démarrant à `startCol` sur `row`. */
  function mergedBlockEndColumn(ws, row, startCol) {
    const cell = ws.getCell(row, startCol);
    if (!cell.isMerged) return startCol;
    const merges = (ws.model && ws.model.merges) || [];
    let end = startCol;
    merges.forEach((m) => {
      const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m);
      if (!match) return;
      const col1 = colNumOf(match[1]), row1 = Number(match[2]), col2 = colNumOf(match[3]), row2 = Number(match[4]);
      if (row >= row1 && row <= row2 && startCol >= col1 && startCol <= col2) end = Math.max(end, col2);
    });
    return end;
  }

  /** Déplace tout un bloc de colonnes [fromStart..fromEnd] vers [toStart..], sur toutes les
   *  lignes de la feuille : valeurs, styles, formats et fusions (jamais les formules elles-mêmes,
   *  puisqu'aucune bibliothèque ExcelJS/openpyxl ne réécrit les références de colonnes lors d'un
   *  déplacement — sans risque ici car ce bloc, tel qu'observé sur les fichiers réels, ne contient
   *  que des valeurs statiques, jamais de formule). Utilisé pour repousser une colonne existante
   *  (ex. "Mouvements en cours") et faire de la place aux 2 nouvelles colonnes. */
  function shiftColumnBlock(ws, fromStart, fromEnd, toStart) {
    if (toStart === fromStart) return;
    const width = fromEnd - fromStart;
    const maxRow = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);

    const merges = ((ws.model && ws.model.merges) || []).slice();
    const movedMerges = [];
    merges.forEach((m) => {
      const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m);
      if (!match) return;
      const col1 = colNumOf(match[1]), row1 = Number(match[2]), col2 = colNumOf(match[3]), row2 = Number(match[4]);
      if (col1 >= fromStart && col2 <= fromEnd) movedMerges.push({ oldRange: m, row1, row2, col1, col2 });
    });
    movedMerges.forEach(({ oldRange }) => ws.unMergeCells(oldRange));

    for (let r = 1; r <= maxRow; r++) {
      for (let i = 0; i <= width; i++) {
        const src = ws.getCell(r, fromStart + i);
        const val = src.value;
        const style = JSON.parse(JSON.stringify(src.style));
        const numFmt = src.numFmt;
        const dst = ws.getCell(r, toStart + i);
        dst.value = val;
        dst.style = style;
        if (numFmt) dst.numFmt = numFmt;
        src.value = null;
        src.style = {};
      }
    }

    for (let i = 0; i <= width; i++) {
      const oldCol = ws.getColumn(fromStart + i);
      const newCol = ws.getColumn(toStart + i);
      if (oldCol.width) newCol.width = oldCol.width;
    }

    const shift = toStart - fromStart;
    movedMerges.forEach(({ row1, row2, col1, col2 }) => {
      ws.mergeCells(row1, col1 + shift, row2, col2 + shift);
    });
  }

  /** Vrai si l'en-tête de cette colonne occupe 2 lignes fusionnées (headerRow:headerRow+1),
   *  comme "Support"/"Dernière Valeur Liquidative" dans les fichiers Althos — pour aligner les 2
   *  nouvelles colonnes sur le même bleu marine "en hauteur" que leurs voisines. */
  function hasTwoRowHeader(ws, headerRow, col) {
    const top = ws.getCell(headerRow, col);
    const bottom = ws.getCell(headerRow + 1, col);
    return !!(top.isMerged && bottom.isMerged && top.master && bottom.master && top.master.address === bottom.master.address);
  }

  /** Étend la zone d'impression déjà définie sur `ws` (le contour bleu visible dans Excel) pour
   *  couvrir au minimum jusqu'à la colonne `minColNeeded` — sans jamais la réduire. Ne fait rien
   *  si aucune zone d'impression n'est définie sur cette feuille. */
  function extendPrintArea(ws, minColNeeded) {
    const area = ws.pageSetup.printArea;
    if (!area) return;
    const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(area.split("!").pop());
    if (!match) return;
    const minCol = colNumOf(match[1]);
    const minRow = Number(match[2]);
    const maxCol = colNumOf(match[3]);
    const maxRow = Number(match[4]);
    if (minColNeeded > maxCol) {
      ws.pageSetup.printArea = `${colLetter(minCol)}${minRow}:${colLetter(minColNeeded)}${maxRow}`;
    }
  }

  /** Étend vers la droite la fusion de cellules qui commence à (row, 1) — typiquement le bandeau
   *  de titre de Consolidation — pour couvrir au minimum jusqu'à `minColNeeded`, sans jamais la
   *  réduire. Ne fait rien si cette cellule n'est pas fusionnée. */
  function extendMergeRight(ws, row, minColNeeded) {
    const cell = ws.getCell(row, 1);
    if (!cell.isMerged) return;
    const end = mergedBlockEndColumn(ws, row, 1);
    if (minColNeeded > end) {
      ws.unMergeCells(row, 1, row, end);
      ws.mergeCells(row, 1, row, minColNeeded);
    }
  }

  function addConsolidationColumns(srcWs, headerRow, totalCol, exitSheetName, firstDataRow, lastDataRow, fundsByIsin) {
    if (!lastDataRow || lastDataRow < firstDataRow) return; // aucun fonds retenu, rien à référencer

    // "Mouvements en cours" (si présente) est toujours la toute dernière colonne du tableau
    // Consolidation, mais n'en fait pas vraiment partie : les 2 nouvelles colonnes doivent être
    // collées juste après la dernière colonne réellement utile (ex. "Dernière Valeur
    // Liquidative"), et "Mouvements en cours" repoussée de 2 colonnes vides après elles.
    let cashCol, penCol;
    const movCol = findColumnByHeaderText(srcWs, headerRow, "Mouvements en cours");
    if (movCol) {
      let coreLast = movCol - 1;
      while (coreLast >= 1 && (srcWs.getCell(headerRow, coreLast).value === null || srcWs.getCell(headerRow, coreLast).value === undefined || String(srcWs.getCell(headerRow, coreLast).value).trim() === "")) {
        coreLast -= 1;
      }
      cashCol = coreLast + 1;
      penCol = coreLast + 2;
      const movEnd = mergedBlockEndColumn(srcWs, headerRow, movCol);
      const desiredMovStart = coreLast + 5; // cashCol, penCol, 2 colonnes vides, puis Mouvements en cours
      if (desiredMovStart > movCol) shiftColumnBlock(srcWs, movCol, movEnd, desiredMovStart);
    } else {
      const lastCol = findLastHeaderColumn(srcWs, headerRow);
      cashCol = nextSafeColumn(srcWs, headerRow, lastCol + 1);
      penCol = nextSafeColumn(srcWs, headerRow, cashCol + 1);
    }

    const headerStyle = cloneStyle(srcWs.getCell(headerRow, 1));
    const baseFontName = (headerStyle.font && headerStyle.font.name) || "Calibri";
    const baseFontSize = (headerStyle.font && headerStyle.font.size) || 10;
    const dataFont = { name: baseFontName, size: baseFontSize, bold: false };
    const boldDataFont = { name: baseFontName, size: baseFontSize, bold: true };
    const twoRowHeader = hasTwoRowHeader(srcWs, headerRow, 1);
    const gridBorder = buildGridBorder(srcWs, headerRow);
    const PENALTY_COL_WIDTH = 63; // même largeur que sur "Calendrier de sortie", pour un rendu cohérent
    const CHARS_PER_LINE = Math.round(PENALTY_COL_WIDTH * 0.55); // cf. commentaire dans buildExitSheet

    [[cashCol, "Rachat — cash reçu", 16], [penCol, "Pénalité de sortie", PENALTY_COL_WIDTH]].forEach(([col, label, width]) => {
      const cell = srcWs.getCell(headerRow, col);
      cell.value = label;
      cell.style = JSON.parse(JSON.stringify(headerStyle));
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = gridBorder;
      const column = srcWs.getColumn(col);
      column.width = width;
      // Toujours visible : la colonne peut hériter un état masqué du fichier d'origine si sa
      // position correspondait déjà à une colonne cachée (ex. l'ancien "Mouvements en cours"
      // déplacé) sans qu'on l'ait explicitement remise à zéro.
      column.hidden = false;
      if (twoRowHeader) {
        const bottomCell = srcWs.getCell(headerRow + 1, col);
        bottomCell.style = JSON.parse(JSON.stringify(headerStyle));
        bottomCell.border = gridBorder;
        srcWs.mergeCells(headerRow, col, headerRow + 1, col);
      }
    });

    const isinRange = `'${exitSheetName}'!$B$${firstDataRow}:$B$${lastDataRow}`;
    const cashRange = `'${exitSheetName}'!$H$${firstDataRow}:$H$${lastDataRow}`;
    const penRange = `'${exitSheetName}'!$I$${firstDataRow}:$I$${lastDataRow}`;

    // Style de bandeau de catégorie (fond beige) : repris de la colonne A de la 1re catégorie
    // trouvée, pour prolonger ce même bandeau sur les 2 nouvelles colonnes plutôt que de laisser
    // un "trou" blanc à chaque ligne de catégorie.
    const { categoryRows, fundRows } = classifyRows(srcWs, headerRow, totalCol);
    const categoryFill = categoryRows.length ? JSON.parse(JSON.stringify(srcWs.getCell(categoryRows[0], 1).style.fill || {})) : null;

    // cashCol/penCol occupent une position de colonne qui existait déjà dans le fichier d'origine
    // (juste après la dernière colonne utile, ou l'ancien emplacement de "Mouvements en cours") :
    // ses cellules peuvent donc porter un fond/quadrillage hérité du fichier client (ex. un
    // second petit tableau récapitulatif de répartition, hors du tableau principal, plus bas sur
    // la feuille). On repart d'une ardoise vierge sur toute la hauteur avant de ne redessiner que
    // les lignes catégorie/fonds du VRAI tableau, pour ne jamais laisser un bloc beige résiduel
    // sans quadrillage.
    const lastSheetRow = srcWs.actualRowCount || srcWs.rowCount;
    const clearFromRow = headerRow + (twoRowHeader ? 2 : 1); // ne jamais effacer la 2e ligne d'un en-tête fusionné sur 2 lignes
    for (let r = clearFromRow; r <= lastSheetRow; r++) {
      [cashCol, penCol].forEach((col) => {
        const cell = srcWs.getCell(r, col);
        // Réassigne l'objet `.style` en entier (pas `cell.fill = ...` isolément) : ExcelJS peut
        // faire partager le même objet de style, en interne, par plusieurs cellules ayant un
        // style identique au chargement — muter une seule propriété dessus contaminerait alors
        // silencieusement d'autres cellules (cf. le commentaire de cloneStyle plus haut).
        cell.style = { fill: { type: "pattern", pattern: "none" }, border: {} };
      });
    }

    // Le quadrillage doit courir sans interruption sur TOUTE la hauteur du VRAI tableau
    // (catégories ET fonds, détenus ou non) — sinon chaque fonds non détenu par ce client laisse
    // un "trou" dans les 2 nouvelles colonnes, puisque Consolidation liste l'univers complet des
    // fonds, pas seulement ceux de ce client.
    // Réassigne `.style` en entier ici aussi (même raison que la boucle de nettoyage ci-dessus) :
    // sans ça, une ligne de fonds pourrait hériter par contamination le fond beige d'une ligne de
    // catégorie voisine avec laquelle ExcelJS aurait fait partager un même objet de style au
    // chargement du fichier d'origine.
    categoryRows.forEach((r) => {
      [cashCol, penCol].forEach((col) => {
        const cell = srcWs.getCell(r, col);
        cell.style = {
          fill: categoryFill ? JSON.parse(JSON.stringify(categoryFill)) : { type: "pattern", pattern: "none" },
          border: gridBorder,
        };
      });
    });
    fundRows.forEach((r) => {
      [cashCol, penCol].forEach((col) => {
        const cell = srcWs.getCell(r, col);
        cell.style = { fill: { type: "pattern", pattern: "none" }, border: gridBorder };
      });
    });

    fundRows.forEach((r) => {
      const isinRaw = srcWs.getCell(r, 2).value;
      const isin = typeof isinRaw === "string" ? isinRaw.trim() : isinRaw;
      if (!isin) return;
      const amount = holdingAmount(srcWs, r, totalCol);
      const isHeld = amount === null || Math.abs(amount) > 0.005;
      if (!isHeld) return;

      const b = `"${isin}"`;
      setF(srcWs, `${colLetter(cashCol)}${r}`, `IFERROR(INDEX(${cashRange},MATCH(${b},${isinRange},0)),"")`);
      const cashCell = srcWs.getCell(r, cashCol);
      cashCell.numFmt = "dd/mm/yyyy";
      cashCell.font = boldDataFont;
      cashCell.alignment = { horizontal: "center", vertical: "middle" };

      setF(srcWs, `${colLetter(penCol)}${r}`, `IFERROR(INDEX(${penRange},MATCH(${b},${isinRange},0)),"")`);
      const penCell = srcWs.getCell(r, penCol);
      // Une formule Excel ne peut jamais renvoyer un texte enrichi (gras partiel) : tout ce
      // qu'affiche cette cellule (message de pénalité ou de fermeture) est mis en gras en bloc.
      penCell.font = boldDataFont;
      penCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

      // Hauteur de 60pt par défaut, augmentée seulement si le texte réel de la base pour ce
      // fonds précis est inhabituellement long — remplace toute hauteur figée héritée du fichier
      // d'origine du conseiller (potentiellement trop courte pour ce nouveau texte).
      const fund = fundsByIsin.get(isin);
      const penInfo = (fund && fund.penalite) || { kind: "inconnue", raw: null };
      const lines = estimatePenaltyLines(penInfo.kind, (penInfo.raw || "").length, (penInfo.dureeVie || "").length, CHARS_PER_LINE);
      srcWs.getRow(r).height = Math.max(DEFAULT_PENALTY_ROW_HEIGHT, heightForLines(lines, baseFontSize));
    });

    extendMergeRight(srcWs, 1, penCol);
    extendPrintArea(srcWs, penCol);
  }

  /** Zone d'impression = tout le tableau, mise à l'échelle sur une page en largeur, paysage —
   *  comme sur la feuille Consolidation. */
  function applyPrintSetup(ws, lastRow, lastVisibleCol) {
    // Marge de 1 ligne/colonne entre le contenu et le contour de la zone d'impression. Pour
    // les colonnes, "le contenu" inclut les colonnes d'aide masquées (jusqu'à HELPER_FIRST_COL +
    // HELPER_NAMES.length - 1), pas seulement les colonnes visibles : sinon le contour passerait
    // au milieu de colonnes masquées au lieu de se terminer 1 colonne après elles.
    const lastHelperCol = HELPER_FIRST_COL + HELPER_NAMES.length - 1;
    const lastCol = Math.max(lastVisibleCol, lastHelperCol);
    ws.pageSetup.printArea = `A1:${colLetter(lastCol + 1)}${lastRow + 1}`;
    ws.pageSetup.orientation = "landscape";
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = 0;
  }

  function setF(ws, addr, formula) {
    ws.getCell(addr).value = { formula };
  }

  function colNumOf(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  // -------------------------------------------------------------------------
  // Point d'entrée
  // -------------------------------------------------------------------------

  async function buildExitCalendarWorkbook(arrayBuffer, fundsData) {
    const FUNDS = fundsData.funds;
    const CALENDAR = fundsData.calendar;
    const fundsByIsin = new Map();
    FUNDS.forEach((f) => { if (f.isin) fundsByIsin.set(f.isin, f); });

    const workbook = new global.ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    // Force Excel à recalculer TOUTES les formules à l'ouverture (le fichier n'a pas de
    // calcChain.xml puisqu'aucune des deux bibliothèques utilisées ici n'exécute les formules
    // elle-même) : évite les cellules calculées qui restent vides tant que l'utilisateur n'a pas
    // forcé un recalcul manuel (Ctrl+Alt+F9).
    workbook.calcProperties = { fullCalcOnLoad: true };

    const srcWs = findConsolidationSheet(workbook);
    if (!srcWs) {
      throw new Error(
        "Feuille « Consolidation » introuvable dans ce classeur (ou aucune feuille ne " +
        "comporte l'en-tête « Support » attendu). Vérifiez qu'il s'agit bien d'un fichier " +
        "de consolidation Althos."
      );
    }
    const headerRow = findHeaderRow(srcWs);
    if (!headerRow) {
      throw new Error("En-tête « Support » introuvable dans les 15 premières lignes de la feuille Consolidation.");
    }

    const wsCal = writeBddCalendrier(workbook, CALENDAR);
    const wsPen = writeBddPenalites(workbook, FUNDS);

    const { ws: exitWs, selectedCount, fundRowsScanned, rowsNeedingDate, firstDataRow, lastDataRow } =
      buildExitSheet(workbook, srcWs, headerRow, CALENDAR, wsCal.rowCount, wsPen.rowCount, fundsByIsin);

    // Complète aussi Consolidation elle-même avec 2 colonnes (cash reçu / pénalité de sortie),
    // en formule vers "Calendrier de sortie" — même information, une seule source de vérité.
    const totalCol = findTotalColumn(srcWs, headerRow);
    addConsolidationColumns(srcWs, headerRow, totalCol, exitWs.name, firstDataRow, lastDataRow, fundsByIsin);

    // Ordre des feuilles : Consolidation, Calendrier de sortie, puis le reste tel quel.
    let order = 1;
    workbook.worksheets.forEach((ws) => {
      if (ws === srcWs) return;
      if (ws === exitWs) return;
      ws.orderNo = 100 + order++;
    });
    srcWs.orderNo = 1;
    exitWs.orderNo = 2;
    workbook.views = [{ activeTab: workbook.worksheets.indexOf(exitWs) }];

    // Le classeur n'est volontairement pas encore sérialisé ici : si rowsNeedingDate n'est pas
    // vide, l'appelant doit d'abord proposer au conseiller de saisir ces dates (directement dans
    // le navigateur, sans repasser par Excel), les injecter dans exitWs colonne D, puis appeler
    // finalizeExitCalendarWorkbook(). S'il n'y a rien à saisir, l'appelant peut finaliser tout de
    // suite.
    return {
      workbook,
      exitWs,
      stats: { selectedCount, fundRowsScanned, headerRow },
      rowsNeedingDate,
    };
  }

  /** Sérialise le classeur en .xlsx (buffer), une fois les éventuelles dates d'investissement
   *  injectées dans exitWs. Étape séparée de buildExitCalendarWorkbook pour permettre au
   *  conseiller de saisir ces dates dans le navigateur avant de générer le fichier final. */
  async function finalizeExitCalendarWorkbook(workbook) {
    const buffer = await workbook.xlsx.writeBuffer();
    return { buffer };
  }

  global.buildExitCalendarWorkbook = buildExitCalendarWorkbook;
  global.finalizeExitCalendarWorkbook = finalizeExitCalendarWorkbook;
})(window);
