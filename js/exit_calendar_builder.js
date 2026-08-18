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
 *   Date d'investissement | Rachat — ordre avant | Rachat — VL | Rachat — exécuté |
 *   Rachat — publié | Rachat — cash reçu | Pénalité de sortie
 *
 * La colonne Pénalité de sortie reste VIDE quand le client n'est plus concerné (délai de
 * pénalité dépassé) ; elle affiche un message dans tous les autres cas (aucune pénalité prévue,
 * pénalité en cours, information non renseignée, ou cas ambigu à vérifier manuellement).
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

  const YELLOW_ARGB = "FFFFF2A6";

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

  /** Classe chaque ligne : 'category' (bandeau, à ignorer), 'blank', ou 'fund'. */
  function classifyRows(ws, headerRow) {
    const categoryRows = [];
    const fundRows = [];
    const lastRow = ws.actualRowCount || ws.rowCount;
    // headerRow+1 = ligne de total général (pas une catégorie, pas un fonds) -> ignorée
    for (let r = headerRow + 2; r <= lastRow; r++) {
      const a = ws.getCell(r, 1);
      if (a.value === null || a.value === undefined || a.value === "") continue;
      if (isSolidBold(a)) {
        categoryRows.push(r);
      } else {
        fundRows.push(r);
      }
    }
    return { categoryRows, fundRows };
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
  const HELPER_FIRST_COL = 10; // J (visible columns now go up to I)

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

    // 1) Repérage des fonds à retenir : détenus (montant total non nul) ET dotés d'un
    //    calendrier de RACHAT connu dans la base Althos. Tout le reste (fonds cotés, fonds
    //    non cotés non détenus, fonds sans calendrier de sortie) n'apparaît pas dans la feuille.
    const totalCol = findTotalColumn(srcWs, headerRow);
    const { fundRows } = classifyRows(srcWs, headerRow);

    const selected = [];
    fundRows.forEach((srcRow) => {
      const isinRaw = srcWs.getCell(srcRow, 2).value;
      const isin = typeof isinRaw === "string" ? isinRaw.trim() : isinRaw;
      const fund = isin ? fundsByIsin.get(isin) : null;
      if (!fund || !hasRachatCalendar(calendar, isin)) return;
      const amount = holdingAmount(srcWs, srcRow, totalCol);
      const isHeld = amount === null || Math.abs(amount) > 0.005;
      if (!isHeld) return;
      selected.push({
        isin,
        nom: srcWs.getCell(srcRow, 1).value || fund.nom,
        amount,
      });
    });

    // 2) En-tête du tableau (style repris de l'en-tête "Support" de la feuille source)
    const headerStyle = cloneStyle(srcWs.getCell(headerRow, 1));
    const headers = [
      "Fonds", "ISIN", "Date d'investissement",
      "Rachat — ordre avant", "Rachat — VL", "Rachat — exécuté",
      "Rachat — publié", "Rachat — cash reçu", "Pénalité de sortie",
    ];
    headers.forEach((label, i) => {
      const cell = ws.getCell(1, i + 1);
      cell.value = label;
      cell.style = JSON.parse(JSON.stringify(headerStyle));
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    });
    ws.getRow(1).height = 30;
    const widths = [34, 14, 16, 16, 14, 14, 14, 14, 46];
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    HELPER_NAMES.forEach((name) => { ws.getColumn(colNumOf(helperCol(name))).hidden = true; });

    if (!selected.length) {
      ws.getCell(2, 1).value =
        "Aucun fonds avec calendrier de sortie (rachat) connu n'est actuellement détenu par ce " +
        "client (montant total nul, ou fonds hors périmètre \"fonds non cotés suivis\").";
      ws.mergeCells(2, 1, 2, headers.length);
      ws.getCell(2, 1).alignment = { wrapText: true, vertical: "middle" };
      ws.getRow(2).height = 30;
      return { ws, selectedCount: 0, fundRowsScanned: fundRows.length };
    }

    // 3) Une ligne par fonds retenu, avec les formules de calcul
    const plainFont = { name: "Arial", size: 10 };
    selected.forEach((item, idx) => {
      const r = idx + 2;
      ws.getCell(r, 1).value = item.nom;
      ws.getCell(r, 2).value = item.isin;
      [1, 2].forEach((c) => { ws.getCell(r, c).font = plainFont; ws.getCell(r, c).alignment = { vertical: "middle" }; });

      const dateCell = ws.getCell(r, 3);
      dateCell.numFmt = "dd/mm/yyyy";
      dateCell.font = plainFont;
      dateCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_ARGB } };

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
      setF(ws, `${N}${r}`, `IF(${M}${r}="","",IFERROR(INDEX(${cal("E")},MATCH(${b}&"|Rachat|"&TEXT(${M}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);
      setF(ws, `${Nx}${r}`, `IF(${M}${r}="","",IFERROR(INDEX(${cal("F")},MATCH(${b}&"|Rachat|"&TEXT(${M}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);
      setF(ws, `${Np}${r}`, `IF(${M}${r}="","",IFERROR(INDEX(${cal("G")},MATCH(${b}&"|Rachat|"&TEXT(${M}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);
      setF(ws, `${O}${r}`, `IF(${M}${r}="","",IFERROR(INDEX(${cal("H")},MATCH(${b}&"|Rachat|"&TEXT(${M}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);

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
      setF(ws, `${AD1}${r}`, `IF(OR(${c}="",${Q}${r}="FUTUR"),"",_xlfn.IFS(${Q}${r}<${U1}${r},${V1}${r},${Q}${r}<${W1}${r},${X1}${r},${Q}${r}<${Y1}${r},${Z1}${r},${Q}${r}<${AA1}${r},${AB1}${r},TRUE,${AC1}${r}))`);

      // Colonnes visibles D..H : valeurs reprises telles quelles de la base (une par champ).
      [["D", M], ["E", N], ["F", Nx], ["G", Np], ["H", O]].forEach(([col, helper]) => {
        setF(ws, `${col}${r}`, `${helper}${r}`);
        const cell = ws.getCell(`${col}${r}`);
        cell.numFmt = "dd/mm/yyyy";
        cell.font = plainFont;
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });

      // Pénalité de sortie : vide si le client n'est plus concerné (délai dépassé) ; un
      // message dans tous les autres cas (aucune pénalité prévue, en cours, non renseignée,
      // ou ambiguë à vérifier manuellement).
      setF(ws, `I${r}`,
        `_xlfn.IFS(` +
        `${S}${r}="aucune","Aucune pénalité de sortie."&IF(${Tc}${r}<>""," ("&${Tc}${r}&")",""),` +
        `${S}${r}="manuel","⚠️ À VÉRIFIER MANUELLEMENT : "&${Tc}${r},` +
        `${S}${r}="inconnue","Pénalité non renseignée — vérifier la notice / DICI du fonds.",` +
        `${c}="","Saisir une date d'investissement pour statuer sur la pénalité.",` +
        `${Q}${r}="FUTUR","Date d'investissement postérieure à aujourd'hui — vérifier la saisie.",` +
        `${AD1}${r}>0,"⚠️ CONCERNÉ : pénalité de "&${AD1}${r}&"% (détention "&${Q}${r}&" mois). "&${Tc}${r},` +
        `TRUE,"")`
      );
      const penCell = ws.getCell(`I${r}`);
      penCell.font = plainFont;
      penCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };

      ws.getRow(r).height = 45;
    });

    return { ws, selectedCount: selected.length, fundRowsScanned: fundRows.length };
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

    const { ws: exitWs, selectedCount, fundRowsScanned } =
      buildExitSheet(workbook, srcWs, headerRow, CALENDAR, wsCal.rowCount, wsPen.rowCount, fundsByIsin);

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

    const buffer = await workbook.xlsx.writeBuffer();
    return {
      buffer,
      stats: { selectedCount, fundRowsScanned, headerRow },
    };
  }

  global.buildExitCalendarWorkbook = buildExitCalendarWorkbook;
})(window);
