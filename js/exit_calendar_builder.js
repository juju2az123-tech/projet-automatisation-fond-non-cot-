/*
 * Ajoute, dans un classeur de consolidation client (feuille "Consolidation" : un fonds par
 * ligne, catégories en lignes surlignées, ISIN en colonne B), une nouvelle feuille
 * "Calendrier de sortie" au même format, avec 5 colonnes pilotées par formules Excel :
 *   Date d'investissement | Prochain ordre — entrée | Prochain ordre — sortie |
 *   Prochaine réception du cash | Pénalité de sortie
 *
 * Port JavaScript (ExcelJS, exécuté dans le navigateur) de scripts/build_client_workbook.py —
 * même logique de calcul, mêmes formules. Contrairement au script Python, la structure du
 * fichier (numéros de ligne des catégories, nombre de colonnes "Contrat") N'EST PAS supposée
 * fixe : elle est détectée dynamiquement pour s'adapter à n'importe quel classeur client réel.
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

  function parseRangeCols(rangeStr) {
    const m = rangeStr.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/);
    if (!m) return null;
    const colNum = (letters) => {
      let n = 0;
      for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
      return n;
    };
    return { c1: colNum(m[1]), r1: +m[2], c2: colNum(m[3]), r2: +m[4] };
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

  // -------------------------------------------------------------------------
  // Construction de la feuille "Calendrier de sortie"
  // -------------------------------------------------------------------------

  const HELPER_NAMES = [
    "has_entree", "next_cutoff_entree", "next_val_entree", "max_cutoff_entree",
    "has_sortie", "next_cutoff_sortie", "next_val_sortie", "next_cash_sortie", "max_cutoff_sortie",
    "months_held", "pen_found", "kind", "raw",
    "max1", "rate1", "max2", "rate2", "max3", "rate3", "max4", "rate4", "rate5",
    "rate_now",
  ];
  const HELPER_FIRST_COL = 8; // H

  function helperCol(name) {
    return colLetter(HELPER_FIRST_COL + HELPER_NAMES.indexOf(name));
  }

  function unmergeOverlapping(ws, fromCol) {
    const merges = ws.model.merges ? ws.model.merges.slice() : [];
    merges.forEach((rangeStr) => {
      const parsed = parseRangeCols(rangeStr);
      if (parsed && parsed.c2 >= fromCol) {
        try { ws.unMergeCells(rangeStr); } catch (e) { /* already gone */ }
      }
    });
  }

  function buildExitSheet(workbook, srcWs, headerRow, calLastRow, penLastRow, fundsByIsin) {
    const cal = (col) => `BDD_Calendrier!$${col}$2:$${col}$${calLastRow}`;
    const pen = (col) => `BDD_Penalites!$${col}$2:$${col}$${penLastRow}`;

    const ws = workbook.addWorksheet("Calendrier de sortie");
    const lastSrcCol = srcWs.actualColumnCount || srcWs.columnCount;
    const lastSrcRow = srcWs.actualRowCount || srcWs.rowCount;

    // 1) Copie intégrale de la feuille source (valeurs + styles + largeurs + fusions),
    //    à l'exception des colonnes C.. (Contrat/TOTAL/etc.), qui seront reconstruites.
    for (let c = 1; c <= lastSrcCol; c++) {
      const col = srcWs.getColumn(c);
      ws.getColumn(c).width = col.width;
    }
    for (let r = 1; r <= lastSrcRow; r++) {
      const srcRow = srcWs.getRow(r);
      const dstRow = ws.getRow(r);
      dstRow.height = srcRow.height;
      for (let c = 1; c <= lastSrcCol; c++) {
        const srcCell = srcWs.getCell(r, c);
        const dstCell = ws.getCell(r, c);
        dstCell.value = srcCell.value && srcCell.value.formula ? null : srcCell.value;
        dstCell.style = cloneStyle(srcCell);
      }
    }
    (srcWs.model.merges || []).forEach((rangeStr) => {
      try { ws.mergeCells(rangeStr); } catch (e) { /* ignore duplicate */ }
    });

    // 2) Purge des colonnes C..fin (Contrat n / TOTAL / Risque / SRI / VL / Mouvements)
    unmergeOverlapping(ws, 3);
    ws.spliceColumns(3, lastSrcCol - 2);

    // 3) En-têtes des 5 nouvelles colonnes, avec le style de l'en-tête d'origine (colonne B)
    const headerStyle = cloneStyle(srcWs.getCell(headerRow, 2));
    const headers = [
      "Date d'investissement",
      "Prochain ordre — entrée",
      "Prochain ordre — sortie",
      "Prochaine réception du cash",
      "Pénalité de sortie",
    ];
    headers.forEach((label, i) => {
      const cell = ws.getCell(headerRow, 3 + i);
      cell.value = label;
      cell.style = JSON.parse(JSON.stringify(headerStyle));
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    });
    ws.getRow(headerRow).height = 42;

    const widths = [16, 32, 32, 30, 46];
    widths.forEach((w, i) => { ws.getColumn(3 + i).width = w; });

    HELPER_NAMES.forEach((name) => { ws.getColumn(colNumOf(helperCol(name))).hidden = true; });

    // 4) Classement des lignes + écriture des formules
    const { categoryRows, fundRows } = classifyRows(ws, headerRow);

    categoryRows.forEach((r) => {
      const bandStyle = cloneStyle(ws.getCell(r, 1));
      for (let c = 3; c <= 7; c++) {
        ws.getCell(r, c).style = JSON.parse(JSON.stringify(bandStyle));
      }
    });

    let dataStyleTemplate = null;
    if (fundRows.length) dataStyleTemplate = cloneStyle(ws.getCell(fundRows[0], 1));

    let nonCoteCount = 0;

    fundRows.forEach((r) => {
      const isinRaw = ws.getCell(r, 2).value;
      const isin = typeof isinRaw === "string" ? isinRaw.trim() : isinRaw;
      const fund = isin ? fundsByIsin.get(isin) : null;

      const dateCell = ws.getCell(r, 3);
      dateCell.style = dataStyleTemplate ? JSON.parse(JSON.stringify(dataStyleTemplate)) : {};
      dateCell.numFmt = "dd/mm/yyyy";

      if (!fund) {
        for (let c = 4; c <= 7; c++) {
          const c2 = ws.getCell(r, c);
          c2.value = "—";
          c2.style = dataStyleTemplate ? JSON.parse(JSON.stringify(dataStyleTemplate)) : {};
          c2.alignment = { horizontal: "center" };
        }
        return;
      }

      nonCoteCount++;
      dateCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: YELLOW_ARGB } };
      ws.getRow(r).height = 60;

      const b = `TRIM($B${r})`;
      const c = `$C${r}`;

      const H = helperCol("has_entree"), I = helperCol("next_cutoff_entree"), J = helperCol("next_val_entree"), K = helperCol("max_cutoff_entree");
      const L = helperCol("has_sortie"), M = helperCol("next_cutoff_sortie"), N = helperCol("next_val_sortie"), O = helperCol("next_cash_sortie"), P = helperCol("max_cutoff_sortie");
      const Q = helperCol("months_held");
      const R = helperCol("pen_found");
      const S = helperCol("kind");
      const Tc = helperCol("raw");
      const U1 = helperCol("max1"), V1 = helperCol("rate1"), W1 = helperCol("max2"), X1 = helperCol("rate2"),
        Y1 = helperCol("max3"), Z1 = helperCol("rate3"), AA1 = helperCol("max4"), AB1 = helperCol("rate4"), AC1 = helperCol("rate5");
      const AD1 = helperCol("rate_now");

      setF(ws, `${H}${r}`, `COUNTIFS(${cal("A")},${b},${cal("C")},"Souscription")`);
      setF(ws, `${I}${r}`, `IF(${H}${r}=0,"",IFERROR(_xlfn.MINIFS(${cal("D")},${cal("A")},${b},${cal("C")},"Souscription",${cal("D")},">="&TODAY()),""))`);
      setF(ws, `${J}${r}`, `IF(${I}${r}="","",IFERROR(INDEX(${cal("E")},MATCH(${b}&"|Souscription|"&TEXT(${I}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);
      setF(ws, `${K}${r}`, `IF(${H}${r}=0,"",_xlfn.MAXIFS(${cal("D")},${cal("A")},${b},${cal("C")},"Souscription"))`);

      setF(ws, `${L}${r}`, `COUNTIFS(${cal("A")},${b},${cal("C")},"Rachat")`);
      setF(ws, `${M}${r}`, `IF(${L}${r}=0,"",IFERROR(_xlfn.MINIFS(${cal("D")},${cal("A")},${b},${cal("C")},"Rachat",${cal("D")},">="&TODAY()),""))`);
      setF(ws, `${N}${r}`, `IF(${M}${r}="","",IFERROR(INDEX(${cal("E")},MATCH(${b}&"|Rachat|"&TEXT(${M}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);
      setF(ws, `${O}${r}`, `IF(${M}${r}="","",IFERROR(INDEX(${cal("H")},MATCH(${b}&"|Rachat|"&TEXT(${M}${r},"yyyy-mm-dd"),${cal("I")},0)),""))`);
      setF(ws, `${P}${r}`, `IF(${L}${r}=0,"",_xlfn.MAXIFS(${cal("D")},${cal("A")},${b},${cal("C")},"Rachat"))`);

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

      setF(ws, `D${r}`, `IF(${H}${r}=0,"Calendrier non disponible pour ce fonds — contacter la société de gestion.",IF(${I}${r}="","Calendrier connu jusqu'au "&TEXT(${K}${r},"dd/mm/yyyy")&" — demander le calendrier à jour.","Cut-off : "&TEXT(${I}${r},"dd/mm/yyyy")&"  |  VL : "&TEXT(${J}${r},"dd/mm/yyyy")))`);
      setF(ws, `E${r}`, `IF(${L}${r}=0,"Calendrier non disponible pour ce fonds — contacter la société de gestion.",IF(${M}${r}="","Calendrier connu jusqu'au "&TEXT(${P}${r},"dd/mm/yyyy")&" — demander le calendrier à jour.","Cut-off : "&TEXT(${M}${r},"dd/mm/yyyy")&"  |  VL : "&TEXT(${N}${r},"dd/mm/yyyy")))`);
      setF(ws, `F${r}`, `IF(${L}${r}=0,"Calendrier non disponible pour ce fonds.",IF(${M}${r}="","Calendrier connu jusqu'au "&TEXT(${P}${r},"dd/mm/yyyy")&".",IF(${O}${r}="","Non précisé dans le calendrier pour cette échéance.","Réception : "&TEXT(${O}${r},"dd/mm/yyyy")&"  (suite au rachat du "&TEXT(${N}${r},"dd/mm/yyyy")&")")))`);
      setF(ws, `G${r}`,
        `_xlfn.IFS(` +
        `${S}${r}="aucune","Aucune pénalité de sortie."&IF(${Tc}${r}<>""," ("&${Tc}${r}&")",""),` +
        `${S}${r}="manuel","⚠️ À VÉRIFIER MANUELLEMENT : "&${Tc}${r},` +
        `${S}${r}="inconnue","Pénalité non renseignée — vérifier la notice / DICI du fonds.",` +
        `${c}="","Saisir une date d'investissement pour statuer sur la pénalité.",` +
        `${Q}${r}="FUTUR","Date d'investissement postérieure à aujourd'hui — vérifier la saisie.",` +
        `${AD1}${r}>0,"⚠️ CONCERNÉ : pénalité de "&${AD1}${r}&"% (détention "&${Q}${r}&" mois). "&${Tc}${r},` +
        `TRUE,"Non concerné (détention "&${Q}${r}&" mois). "&${Tc}${r})`
      );

      for (let c2 = 4; c2 <= 7; c2++) {
        ws.getCell(r, c2).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      }
    });

    return { ws, nonCoteCount, fundRowsCount: fundRows.length, categoryRowsCount: categoryRows.length };
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

    const { ws: exitWs, nonCoteCount, fundRowsCount, categoryRowsCount } =
      buildExitSheet(workbook, srcWs, headerRow, wsCal.rowCount, wsPen.rowCount, fundsByIsin);

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
      stats: { nonCoteCount, fundRowsCount, categoryRowsCount, headerRow },
    };
  }

  global.buildExitCalendarWorkbook = buildExitCalendarWorkbook;
})(window);
