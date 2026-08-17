#!/usr/bin/env python3
"""
Ajoute au classeur de consolidation client (ex : source/ConsolidationTemplateAlthosAI_V5.xlsx)
une nouvelle feuille "Calendrier de sortie", clonée de la feuille "Consolidation" (même
présentation : mêmes lignes de fonds, mêmes catégories, mêmes styles), dans laquelle les
colonnes "Contrat n" / TOTAL / Risque / SRI / VL / Mouvements sont remplacées par 5 colonnes :

  - Date d'investissement (à saisir par le conseiller)
  - Prochain ordre — entrée
  - Prochain ordre — sortie
  - Prochaine réception du cash
  - Pénalité de sortie (statut calculé par rapport à la date d'investissement saisie)

Toutes les colonnes calculées sont des FORMULES Excel (recalculées à chaque ouverture, à la
date du jour), et s'appuient sur deux feuilles de données ajoutées et masquées :
  - BDD_Calendrier   (calendrier mensuel des fonds, dérivé de Calendriers_de_fonds_Althos.xlsx)
  - BDD_Penalites    (règles de pénalité de sortie structurées, même parseur que build_data.py)

Usage :
    python3 scripts/build_client_workbook.py [chemin_consolidation.xlsx] [chemin_sortie.xlsx]

Par défaut lit source/ConsolidationTemplateAlthosAI_V5.xlsx et écrit
output/ConsolidationTemplateAlthosAI_V5_avec_calendrier.xlsx
"""
import sys
import datetime
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_data import load_suivi, merge_extra_from_calendriers_par_fonds, load_calendrier  # noqa: E402

CAL_FILE = ROOT / "source" / "Calendriers_de_fonds_Althos.xlsx"

NON_COTE_ROW_MIN = 100  # 1ère ligne de fonds de la section "Hedge Fund"
NON_COTE_ROW_MAX = 310  # dernière ligne de fonds de la section "Private Equity"
FUND_ROW_FIRST = 8
FUND_ROW_LAST = 460

YELLOW = PatternFill(start_color="FFFFF2A6", end_color="FFFFF2A6", fill_type="solid")


# ---------------------------------------------------------------------------
# Row classification (reuses the same visual convention as the source sheet:
# a "category" row is solid-filled with no ISIN in column B).
# ---------------------------------------------------------------------------

def classify_rows(ws):
    cat_rows, fund_rows, blank_rows = [], [], []
    for r in range(FUND_ROW_FIRST, FUND_ROW_LAST + 1):
        a = ws.cell(row=r, column=1)
        b = ws.cell(row=r, column=2)
        if a.value is None:
            blank_rows.append(r)
        elif a.fill.patternType == "solid" and b.value is None:
            cat_rows.append(r)
        else:
            fund_rows.append(r)
    return cat_rows, fund_rows, blank_rows


# ---------------------------------------------------------------------------
# Pénalité -> 9 valeurs numériques exploitables par une formule IFS en cascade
# (4 paliers "seuil" + 1 taux "au-delà"), quel que soit le nombre réel de
# paliers d'origine (seuil simple, dégressif à N paliers, etc.)
# ---------------------------------------------------------------------------

def build_tier_columns(pen):
    kind = pen["kind"]
    if kind in ("seuil", "soft"):
        t = pen["tiers"][0]
        display = [{"max": t["maxMonths"], "rate": t["rate"]}, {"max": None, "rate": 0}]
    elif kind == "degressif":
        display = [{"max": t["maxMonths"], "rate": t["rate"]} for t in pen["tiers"]]
    else:
        return [0, 0, 0, 0, 0, 0, 0, 0, 0]

    conditions = display[:-1]
    tail_rate = display[-1]["rate"]
    while len(conditions) < 4:
        last_max = conditions[-1]["max"] if conditions else 0
        conditions.append({"max": last_max, "rate": tail_rate})
    conditions = conditions[:4]

    out = []
    for c in conditions:
        out.append(c["max"])
        out.append(c["rate"])
    out.append(tail_rate)
    return out  # [Max1,Rate1,Max2,Rate2,Max3,Rate3,Max4,Rate4,Rate5]


# ---------------------------------------------------------------------------
# BDD_Calendrier
# ---------------------------------------------------------------------------

def write_bdd_calendrier(wb, calendar):
    ws = wb.create_sheet("BDD_Calendrier")
    headers = ["ISIN", "Nom", "Type", "Cutoff", "Valorisation", "Execution",
               "PublicationVL", "ReglementCash", "CleComposite"]
    ws.append(headers)
    row_i = 2
    for isin, by_type in calendar.items():
        for type_, entries in by_type.items():
            expanded_types = ["Souscription", "Rachat"] if type_ == "Souscription et rachat" else [type_]
            for etype in expanded_types:
                for e in entries:
                    cutoff = e["cutoff"]
                    key = f"{isin}|{etype}|{cutoff}" if cutoff else ""
                    ws.append([
                        isin, "", etype,
                        _to_date(e["cutoff"]), _to_date(e["valorisation"]), _to_date(e["execution"]),
                        _to_date(e["publicationVL"]), _to_date(e["reglementCash"]), key,
                    ])
                    row_i += 1
    for col, width in zip("ABCDEFGHI", [14, 4, 12, 12, 12, 12, 12, 12, 26]):
        ws.column_dimensions[col].width = width
    for c in ws[1]:
        c.font = openpyxl.styles.Font(bold=True)
    ws.sheet_state = "hidden"
    return ws


def _to_date(iso):
    if not iso:
        return None
    y, m, d = iso.split("-")
    return datetime.date(int(y), int(m), int(d))


# ---------------------------------------------------------------------------
# BDD_Penalites
# ---------------------------------------------------------------------------

def write_bdd_penalites(wb, funds):
    ws = wb.create_sheet("BDD_Penalites")
    headers = ["ISIN", "Nom", "Kind", "Max1", "Rate1", "Max2", "Rate2",
               "Max3", "Rate3", "Max4", "Rate4", "Rate5", "RawText"]
    ws.append(headers)
    for key, f in funds.items():
        isin = f["isin"]
        if not isin:
            continue
        pen = f["penalite"]
        tiers9 = build_tier_columns(pen)
        ws.append([isin, f["nom"], pen["kind"], *tiers9, pen["raw"] or ""])
    for col, width in zip("ABCDEFGHIJKLM", [14, 34, 11, 7, 7, 7, 7, 7, 7, 7, 7, 7, 60]):
        ws.column_dimensions[col].width = width
    for c in ws[1]:
        c.font = openpyxl.styles.Font(bold=True)
    ws.sheet_state = "hidden"
    return ws


# ---------------------------------------------------------------------------
# Nouvelle feuille visible "Calendrier de sortie"
# ---------------------------------------------------------------------------

HELPER_NAMES = [
    "has_entree", "next_cutoff_entree", "next_val_entree", "max_cutoff_entree",
    "has_sortie", "next_cutoff_sortie", "next_val_sortie", "next_cash_sortie", "max_cutoff_sortie",
    "months_held", "pen_found", "kind", "raw",
    "max1", "rate1", "max2", "rate2", "max3", "rate3", "max4", "rate4", "rate5",
    "rate_now",
]
HELPER_FIRST_COL = 8  # H


def helper_col(name):
    idx = HELPER_FIRST_COL + HELPER_NAMES.index(name)
    return get_column_letter(idx)


def build_exit_sheet(wb, src_ws, cal_last_row, pen_last_row):
    def cal(col):
        return f"BDD_Calendrier!${col}$2:${col}${cal_last_row}"

    def pen(col):
        return f"BDD_Penalites!${col}$2:${col}${pen_last_row}"

    ws = wb.copy_worksheet(src_ws)
    ws.title = "Calendrier de sortie"

    # Move it right after "Consolidation"
    wb._sheets.remove(ws)
    wb._sheets.insert(wb._sheets.index(src_ws) + 1, ws)

    header_style_source = copy(ws["B6"]._style)  # navy header style (theme1, bold, white)
    data_style_template = copy(ws["B101"]._style)  # plain fund-row style (left align, no fill)

    # Remove the old "Contrat.../TOTAL/Risque/SRI/VL/Mouvements" block. openpyxl's
    # delete_cols() does not adjust merged-cell metadata (and errors out later if a
    # stale merge still references a cell it removed), so unmerge everything that
    # overlaps C:AA — including the footnote merge — before deleting the columns.
    for rng in ("T6:U6", "X6:X7", "Z6:AA7", "A461:D461"):
        ws.unmerge_cells(rng)
    ws.delete_cols(3, 25)  # C..AA (25 columns)

    headers = [
        "Date d'investissement",
        "Prochain ordre — entrée",
        "Prochain ordre — sortie",
        "Prochaine réception du cash",
        "Pénalité de sortie",
    ]
    for i, label in enumerate(headers):
        col = 3 + i  # C..G
        cell = ws.cell(row=6, column=col, value=label)
        cell._style = copy(header_style_source)
        cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[6].height = 42

    widths = [16, 32, 32, 30, 46]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(3 + i)].width = w

    # Hide helper columns
    for name in HELPER_NAMES:
        ws.column_dimensions[helper_col(name)].hidden = True

    # Footnote row (previously A461:D461, already unmerged above) -> widen, refresh wording
    ws["A461"] = ("Colonnes calculées automatiquement à l'ouverture du fichier, à partir des "
                  "calendriers officiels et pénalités de sortie Althos. À vérifier contre le "
                  "PDF / DICI du fonds avant toute réponse engageante au client.")
    ws.merge_cells("A461:G461")
    for r in range(462, 468):
        for c in range(1, 3):
            ws.cell(row=r, column=c).value = None

    cat_rows, fund_rows, _ = classify_rows(ws)

    # Re-apply full-row category banding across the new columns C..G
    for r in cat_rows:
        band_style = copy(ws.cell(row=r, column=1)._style)
        for col in range(3, 8):
            cell = ws.cell(row=r, column=col)
            cell._style = band_style

    for r in fund_rows:
        is_non_cote = NON_COTE_ROW_MIN <= r <= NON_COTE_ROW_MAX
        date_cell = ws.cell(row=r, column=3)
        date_cell._style = copy(data_style_template)
        date_cell.number_format = "dd/mm/yyyy"

        if not is_non_cote:
            date_cell.fill = PatternFill(fill_type=None)
            for col in range(4, 8):
                c = ws.cell(row=r, column=col, value="—")
                c._style = copy(data_style_template)
                c.alignment = openpyxl.styles.Alignment(horizontal="center")
            continue

        date_cell.fill = YELLOW
        ws.row_dimensions[r].height = 60  # room for the wrapped multi-line status text
        # TRIM() guards against stray trailing spaces on ISINs in the source template
        # (confirmed present on a few rows, e.g. "FR0013186772 ") that would otherwise
        # silently break the exact-match lookups against BDD_Calendrier/BDD_Penalites.
        b = f"TRIM($B{r})"
        c = f"$C{r}"

        H, I, J, K = helper_col("has_entree"), helper_col("next_cutoff_entree"), helper_col("next_val_entree"), helper_col("max_cutoff_entree")
        L, M, N, O, P = helper_col("has_sortie"), helper_col("next_cutoff_sortie"), helper_col("next_val_sortie"), helper_col("next_cash_sortie"), helper_col("max_cutoff_sortie")
        Q = helper_col("months_held")
        R = helper_col("pen_found")
        S = helper_col("kind")
        Tc = helper_col("raw")
        U1, V1, W1, X1, Y1, Z1, AA1, AB1, AC1 = (helper_col("max1"), helper_col("rate1"), helper_col("max2"),
                                                  helper_col("rate2"), helper_col("max3"), helper_col("rate3"),
                                                  helper_col("max4"), helper_col("rate4"), helper_col("rate5"))
        AD1 = helper_col("rate_now")

        ws[f"{H}{r}"] = f'=COUNTIFS({cal("A")},{b},{cal("C")},"Souscription")'
        ws[f"{I}{r}"] = (f'=IF({H}{r}=0,"",IFERROR(_xlfn.MINIFS({cal("D")},'
                          f'{cal("A")},{b},{cal("C")},"Souscription",'
                          f'{cal("D")},">="&TODAY()),""))')
        ws[f"{J}{r}"] = (f'=IF({I}{r}="","",IFERROR(INDEX({cal("E")},'
                          f'MATCH({b}&"|Souscription|"&TEXT({I}{r},"yyyy-mm-dd"),{cal("I")},0)),""))')
        ws[f"{K}{r}"] = f'=IF({H}{r}=0,"",_xlfn.MAXIFS({cal("D")},{cal("A")},{b},{cal("C")},"Souscription"))'

        ws[f"{L}{r}"] = f'=COUNTIFS({cal("A")},{b},{cal("C")},"Rachat")'
        ws[f"{M}{r}"] = (f'=IF({L}{r}=0,"",IFERROR(_xlfn.MINIFS({cal("D")},'
                          f'{cal("A")},{b},{cal("C")},"Rachat",'
                          f'{cal("D")},">="&TODAY()),""))')
        ws[f"{N}{r}"] = (f'=IF({M}{r}="","",IFERROR(INDEX({cal("E")},'
                          f'MATCH({b}&"|Rachat|"&TEXT({M}{r},"yyyy-mm-dd"),{cal("I")},0)),""))')
        ws[f"{O}{r}"] = (f'=IF({M}{r}="","",IFERROR(INDEX({cal("H")},'
                          f'MATCH({b}&"|Rachat|"&TEXT({M}{r},"yyyy-mm-dd"),{cal("I")},0)),""))')
        ws[f"{P}{r}"] = f'=IF({L}{r}=0,"",_xlfn.MAXIFS({cal("D")},{cal("A")},{b},{cal("C")},"Rachat"))'

        ws[f"{Q}{r}"] = f'=IF({c}="","",IF({c}>TODAY(),"FUTUR",DATEDIF({c},TODAY(),"m")))'
        ws[f"{R}{r}"] = f'=COUNTIF({pen("A")},{b})'
        ws[f"{S}{r}"] = f'=IF({R}{r}=0,"inconnue",INDEX({pen("C")},MATCH({b},{pen("A")},0)))'
        ws[f"{Tc}{r}"] = f'=IF({R}{r}=0,"",INDEX({pen("M")},MATCH({b},{pen("A")},0)))'
        ws[f"{U1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("D")},MATCH({b},{pen("A")},0)))'
        ws[f"{V1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("E")},MATCH({b},{pen("A")},0)))'
        ws[f"{W1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("F")},MATCH({b},{pen("A")},0)))'
        ws[f"{X1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("G")},MATCH({b},{pen("A")},0)))'
        ws[f"{Y1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("H")},MATCH({b},{pen("A")},0)))'
        ws[f"{Z1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("I")},MATCH({b},{pen("A")},0)))'
        ws[f"{AA1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("J")},MATCH({b},{pen("A")},0)))'
        ws[f"{AB1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("K")},MATCH({b},{pen("A")},0)))'
        ws[f"{AC1}{r}"] = f'=IF({R}{r}=0,0,INDEX({pen("L")},MATCH({b},{pen("A")},0)))'
        ws[f"{AD1}{r}"] = (f'=IF(OR({c}="",{Q}{r}="FUTUR"),"",_xlfn.IFS({Q}{r}<{U1}{r},{V1}{r},{Q}{r}<{W1}{r},{X1}{r},'
                            f'{Q}{r}<{Y1}{r},{Z1}{r},{Q}{r}<{AA1}{r},{AB1}{r},TRUE,{AC1}{r}))')

        # Colonnes visibles
        ws[f"D{r}"] = (f'=IF({H}{r}=0,"Calendrier non disponible pour ce fonds — contacter la société de gestion.",'
                        f'IF({I}{r}="","Calendrier connu jusqu\'au "&TEXT({K}{r},"dd/mm/yyyy")&" — demander le calendrier à jour.",'
                        f'"Cut-off : "&TEXT({I}{r},"dd/mm/yyyy")&"  |  VL : "&TEXT({J}{r},"dd/mm/yyyy")))')
        ws[f"E{r}"] = (f'=IF({L}{r}=0,"Calendrier non disponible pour ce fonds — contacter la société de gestion.",'
                        f'IF({M}{r}="","Calendrier connu jusqu\'au "&TEXT({P}{r},"dd/mm/yyyy")&" — demander le calendrier à jour.",'
                        f'"Cut-off : "&TEXT({M}{r},"dd/mm/yyyy")&"  |  VL : "&TEXT({N}{r},"dd/mm/yyyy")))')
        ws[f"F{r}"] = (f'=IF({L}{r}=0,"Calendrier non disponible pour ce fonds.",'
                        f'IF({M}{r}="","Calendrier connu jusqu\'au "&TEXT({P}{r},"dd/mm/yyyy")&".",'
                        f'IF({O}{r}="","Non précisé dans le calendrier pour cette échéance.",'
                        f'"Réception : "&TEXT({O}{r},"dd/mm/yyyy")&"  (suite au rachat du "&TEXT({N}{r},"dd/mm/yyyy")&")")))')
        ws[f"G{r}"] = (
            f'=_xlfn.IFS('
            f'{S}{r}="aucune","Aucune pénalité de sortie."&IF({Tc}{r}<>""," ("&{Tc}{r}&")",""),'
            f'{S}{r}="manuel","⚠️ À VÉRIFIER MANUELLEMENT : "&{Tc}{r},'
            f'{S}{r}="inconnue","Pénalité non renseignée — vérifier la notice / DICI du fonds.",'
            f'{c}="","Saisir une date d\'investissement pour statuer sur la pénalité.",'
            f'{Q}{r}="FUTUR","Date d\'investissement postérieure à aujourd\'hui — vérifier la saisie.",'
            f'{AD1}{r}>0,"⚠️ CONCERNÉ : pénalité de "&{AD1}{r}&"% (détention "&{Q}{r}&" mois). "&{Tc}{r},'
            f'TRUE,"Non concerné (détention "&{Q}{r}&" mois). "&{Tc}{r})'
        )
        for col in range(4, 8):
            ws.cell(row=r, column=col).alignment = openpyxl.styles.Alignment(horizontal="left", vertical="center", wrap_text=True)

    return ws


def main():
    src_path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "source" / "ConsolidationTemplateAlthosAI_V5.xlsx"
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "output" / "ConsolidationTemplateAlthosAI_V5_avec_calendrier.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb_cal = openpyxl.load_workbook(CAL_FILE, data_only=True)
    funds = load_suivi(wb_cal)
    funds = merge_extra_from_calendriers_par_fonds(wb_cal, funds)
    calendar = load_calendrier(wb_cal)

    wb = openpyxl.load_workbook(src_path, data_only=False)
    src_ws = wb["Consolidation"]

    ws_cal = write_bdd_calendrier(wb, calendar)
    ws_pen = write_bdd_penalites(wb, funds)
    build_exit_sheet(wb, src_ws, ws_cal.max_row, ws_pen.max_row)

    wb.active = wb.sheetnames.index("Calendrier de sortie")
    wb.save(out_path)
    print(f"OK -> {out_path}")


if __name__ == "__main__":
    main()
