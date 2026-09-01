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
      "Max3", "Rate3", "Max4", "Rate4", "Rate5", "RawText"];
    ws.addRow(headers);

    funds.forEach((f) => {
      if (!f.isin) return;
      const pen = f.penalite || { kind: "inconnue", raw: null, tiers: [] };
      const tiers9 = buildTierColumns(pen);
      ws.addRow([f.isin, f.nom, pen.kind, ...tiers9, pen.raw || ""]);
    });

    const widths = [14, 34, 11, 7, 7, 7, 7, 7, 7, 7, 7, 7, 60];
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
  function classifyRows(ws, headerRow) {
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
    "months_held", "pen_found", "kind", "raw",
    "max1", "rate1", "max2", "rate2", "max3", "rate3", "max4", "rate4", "rate5",
    "rate_now",
  ];
  const HELPER_FIRST_COL = 11; // K (visible columns now go up to J : Titulaire ajouté en colonne A)

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
    const { fundRows, rowToCategory } = classifyRows(srcWs, headerRow);
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
        selected.push({ isin, nom, category, owner, amount, needsDate });
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
    // Style du bandeau de titulaire : repris tel quel de la ligne de Consolidation où figurent
    // les libellés "Monsieur" / "Madame" / société (juste au-dessus de l'en-tête "Support") —
    // même principe que pour les autres styles de cette feuille : jamais une couleur codée en dur.
    let ownerHeaderStyle = null;
    if (showOwnerHeadings) {
      const ownerRow = headerRow - 1;
      for (let c = 3; c < totalCol; c++) {
        const v = srcWs.getCell(ownerRow, c).value;
        if (v !== null && v !== undefined && String(v).trim() !== "") {
          ownerHeaderStyle = cloneStyle(srcWs.getCell(ownerRow, c));
          break;
        }
      }
    }

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

    const titleStyle = cloneStyle(srcWs.getCell(1, 1));
    const subtitleStyle = cloneStyle(srcWs.getCell(3, 1));
    // Style de bandeau de catégorie : repris tel quel de la première ligne de catégorie trouvée
    // dans la Consolidation (même couleur beige, même police en gras).
    const { categoryRows } = classifyRows(srcWs, headerRow);
    const categoryStyle = categoryRows.length ? cloneStyle(srcWs.getCell(categoryRows[0], 1)) : null;
    // Couleur de quadrillage : reprise telle quelle (référence de thème + teinte, PAS une valeur
    // fixe) de la bordure déjà utilisée par l'en-tête de Consolidation — sa couleur réelle dépend
    // du thème propre à chaque classeur (beige dans certains, bleu dans d'autres), donc on ne
    // peut pas la coder en dur.
    const headerBorder = srcWs.getCell(headerRow, 1).border || {};
    const gridSide = JSON.parse(JSON.stringify(
      headerBorder.left || headerBorder.top || headerBorder.right || headerBorder.bottom ||
      { style: "thin", color: { argb: "FFDDCCB8" } }
    ));
    gridSide.style = "medium";
    const gridBorder = { top: gridSide, left: gridSide, bottom: gridSide, right: gridSide };

    const headers = [
      "Fonds", "ISIN", "Date d'investissement",
      "Rachat — ordre avant", "Rachat — VL", "Rachat — exécuté",
      "Rachat — publié", "Rachat — cash reçu", "Pénalité de sortie",
    ];
    const HEADER_ROW_OUT = 5; // 1: titre, 2: (vide), 3: sous-titre daté, 4: (vide), 5: en-tête
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

    headers.forEach((label, i) => {
      const cell = ws.getCell(HEADER_ROW_OUT, i + 1);
      cell.value = label;
      cell.style = JSON.parse(JSON.stringify(headerStyle));
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: false };
      cell.border = gridBorder;
    });
    ws.getRow(HEADER_ROW_OUT).height = 22;
    const widths = [32, 14, 22, 22, 15, 19, 18, 20, 46];
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

    // 3) Un tableau complètement séparé par titulaire (bandeau de titulaire, puis bandeaux de
    //    catégorie beige, puis une ligne par fonds retenu), avec les formules de calcul.
    let r = FIRST_DATA_ROW;
    const rowsNeedingDate = [];
    ownerOrder.forEach((ownerKey, ownerIdx) => {
      if (showOwnerHeadings) {
        const headingStyle = ownerHeaderStyle || headerStyle;
        ws.getCell(r, 1).value = ownerKey || "Autres titulaires";
        for (let c = 1; c <= headers.length; c++) {
          ws.getCell(r, c).style = JSON.parse(JSON.stringify(headingStyle));
        }
        ws.mergeCells(r, 1, r, headers.length);
        ws.getCell(r, 1).alignment = { horizontal: "center", vertical: "middle" };
        ws.getRow(r).height = 20;
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
          `${S}${r}="ferme","FONDS FERMÉ — sortie non disponible (sauf cas exceptionnel prévu au règlement, ex. décès, invalidité). Vérifier le DICI du fonds.",` +
          `${S}${r}="manuel","À VÉRIFIER MANUELLEMENT : "&${Tc}${r},` +
          `${S}${r}="aucune","",` +
          `${S}${r}="inconnue","",` +
          `${c}="","Saisir une date d'investissement pour statuer sur la pénalité.",` +
          `${Q}${r}="FUTUR","Date d'investissement postérieure à aujourd'hui — vérifier la saisie.",` +
          `${AD1}${r}>0,"CONCERNÉ : pénalité de "&${AD1}${r}&"% (détention "&${Q}${r}&" mois). "&${Tc}${r},` +
          `TRUE(),"")`
        );
        const penCell = ws.getCell(`I${r}`);
        penCell.font = dataFont;
        penCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

        ws.getRow(r).height = 60;
        r += 1;
      });

      // Ligne vide (non bordée) entre 2 titulaires, pour que le contour du tableau ne referme
      // qu'un seul titulaire à la fois, jamais les deux ensemble.
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

  function addConsolidationColumns(srcWs, headerRow, totalCol, exitSheetName, firstDataRow, lastDataRow) {
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
    const twoRowHeader = hasTwoRowHeader(srcWs, headerRow, 1);

    [[cashCol, "Rachat — cash reçu", 16], [penCol, "Pénalité de sortie", 46]].forEach(([col, label, width]) => {
      const cell = srcWs.getCell(headerRow, col);
      cell.value = label;
      cell.style = JSON.parse(JSON.stringify(headerStyle));
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      srcWs.getColumn(col).width = width;
      if (twoRowHeader) {
        const bottomCell = srcWs.getCell(headerRow + 1, col);
        bottomCell.style = JSON.parse(JSON.stringify(headerStyle));
        srcWs.mergeCells(headerRow, col, headerRow + 1, col);
      }
    });

    const isinRange = `'${exitSheetName}'!$B$${firstDataRow}:$B$${lastDataRow}`;
    const cashRange = `'${exitSheetName}'!$H$${firstDataRow}:$H$${lastDataRow}`;
    const penRange = `'${exitSheetName}'!$I$${firstDataRow}:$I$${lastDataRow}`;

    const { fundRows } = classifyRows(srcWs, headerRow);
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
      cashCell.font = dataFont;
      cashCell.alignment = { horizontal: "center", vertical: "middle" };

      setF(srcWs, `${colLetter(penCol)}${r}`, `IFERROR(INDEX(${penRange},MATCH(${b},${isinRange},0)),"")`);
      const penCell = srcWs.getCell(r, penCol);
      penCell.font = dataFont;
      penCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    });
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
    addConsolidationColumns(srcWs, headerRow, totalCol, exitWs.name, firstDataRow, lastDataRow);

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
