#!/usr/bin/env python3
"""
Ajoute au classeur de consolidation client (ex : source/ConsolidationTemplateAlthosAI_V5.xlsx)
une nouvelle feuille "Calendrier de sortie" — un tableau COMPACT listant uniquement les fonds
que le client détient réellement (montant total non nul dans sa consolidation) ET pour lesquels
un calendrier de sortie officiel est connu dans la base Althos. Ni les fonds cotés, ni les fonds
non détenus, ni les fonds non cotés sans calendrier n'apparaissent : ce n'est pas une copie de la
feuille "Consolidation", juste la liste utile pour répondre à un client qui demande ses délais
de sortie. Pour chaque fonds retenu, 5 colonnes :

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
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_data import (  # noqa: E402
    load_suivi, merge_extra_from_calendriers_par_fonds, load_calendrier,
    load_calendrier_par_fonds, merge_calendriers,
)

CAL_FILE = ROOT / "source" / "Calendriers_de_fonds_Althos.xlsx"

YELLOW = PatternFill(start_color="FFFFF2A6", end_color="FFFFF2A6", fill_type="solid")
PLAIN_FONT = Font(name="Arial", size=10)


# ---------------------------------------------------------------------------
# Détection dynamique de la structure de la feuille "Consolidation" (aucun
# numéro de ligne/colonne n'est supposé fixe, pour s'adapter à un fichier
# client réel dont la mise en page peut différer du template).
# ---------------------------------------------------------------------------

def find_header_row(ws):
    for r in range(1, 16):
        if (ws.cell(row=r, column=1).value or "").strip() == "Support":
            return r
    return None


def find_total_column(ws, header_row):
    for c in range(3, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v and str(v).strip().upper() == "TOTAL":
            return c
    return None


def classify_rows(ws, header_row):
    """Une ligne de catégorie (bandeau) est en gras + fond uni, sans ISIN en colonne B."""
    cat_rows, fund_rows = [], []
    for r in range(header_row + 2, ws.max_row + 1):  # header_row+1 = ligne de total général
        a = ws.cell(row=r, column=1)
        b = ws.cell(row=r, column=2)
        if a.value is None:
            continue
        if a.fill.patternType == "solid" and (a.font.bold or False) and b.value is None:
            cat_rows.append(r)
        else:
            fund_rows.append(r)
    return cat_rows, fund_rows


def holding_amount(ws, row, total_col):
    """Somme des colonnes 'Contrat' (3..total_col-1) ; None si aucune colonne TOTAL trouvée."""
    if not total_col:
        return None
    total = 0
    for c in range(3, total_col):
        v = ws.cell(row=row, column=c).value
        if isinstance(v, (int, float)):
            total += v
    return total


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


def build_exit_sheet(wb, src_ws, cal_last_row, pen_last_row, funds_by_isin):
    def cal(col):
        return f"BDD_Calendrier!${col}$2:${col}${cal_last_row}"

    def pen(col):
        return f"BDD_Penalites!${col}$2:${col}${pen_last_row}"

    ws = wb.create_sheet("Calendrier de sortie")
    wb._sheets.remove(ws)
    wb._sheets.insert(wb._sheets.index(src_ws) + 1, ws)  # right after "Consolidation"

    header_row = find_header_row(src_ws)
    if header_row is None:
        raise ValueError('En-tête "Support" introuvable dans les 15 premières lignes de la feuille Consolidation.')
    total_col = find_total_column(src_ws, header_row)
    _, fund_rows = classify_rows(src_ws, header_row)

    selected = []
    for r in fund_rows:
        isin_raw = src_ws.cell(row=r, column=2).value
        isin = isin_raw.strip() if isinstance(isin_raw, str) else isin_raw
        fund = funds_by_isin.get(isin) if isin else None
        if not fund or not fund.get("hasCalendar"):
            continue
        amount = holding_amount(src_ws, r, total_col)
        if amount is not None and abs(amount) <= 0.005:
            continue
        selected.append((isin, src_ws.cell(row=r, column=1).value or fund["nom"]))

    header_style = copy(src_ws.cell(row=header_row, column=1)._style)  # navy header (theme1)
    headers = ["Fonds", "ISIN", "Date d'investissement",
               "Prochain ordre — entrée", "Prochain ordre — sortie",
               "Prochaine réception du cash", "Pénalité de sortie"]
    for i, label in enumerate(headers):
        cell = ws.cell(row=1, column=i + 1, value=label)
        cell._style = copy(header_style)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    widths = [34, 14, 16, 32, 32, 30, 46]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    for name in HELPER_NAMES:
        ws.column_dimensions[helper_col(name)].hidden = True

    if not selected:
        msg = ("Aucun fonds avec calendrier de sortie connu n'est actuellement détenu par ce "
               'client (montant total nul, ou fonds hors périmètre "fonds non cotés suivis").')
        ws.cell(row=2, column=1, value=msg)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        ws.cell(row=2, column=1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[2].height = 30
        return ws, 0, len(fund_rows)

    for idx, (isin, nom) in enumerate(selected):
        r = idx + 2
        ws.cell(row=r, column=1, value=nom).font = PLAIN_FONT
        ws.cell(row=r, column=2, value=isin).font = PLAIN_FONT

        date_cell = ws.cell(row=r, column=3)
        date_cell.number_format = "dd/mm/yyyy"
        date_cell.font = PLAIN_FONT
        date_cell.fill = YELLOW
        ws.row_dimensions[r].height = 60  # room for the wrapped multi-line status text

        b = f'"{isin}"'
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
            cell = ws.cell(row=r, column=col)
            cell.font = PLAIN_FONT
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    return ws, len(selected), len(fund_rows)


def main():
    src_path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "source" / "ConsolidationTemplateAlthosAI_V5.xlsx"
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "output" / "ConsolidationTemplateAlthosAI_V5_avec_calendrier.xlsx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb_cal = openpyxl.load_workbook(CAL_FILE, data_only=True)
    funds = load_suivi(wb_cal)
    funds = merge_extra_from_calendriers_par_fonds(wb_cal, funds)
    calendar = load_calendrier(wb_cal)
    calendar = merge_calendriers(calendar, load_calendrier_par_fonds(wb_cal))
    for f in funds.values():
        f["hasCalendar"] = bool(f["isin"] and f["isin"] in calendar)
    funds_by_isin = {f["isin"]: f for f in funds.values() if f["isin"]}

    wb = openpyxl.load_workbook(src_path, data_only=False)
    src_ws = wb["Consolidation"]

    ws_cal = write_bdd_calendrier(wb, calendar)
    ws_pen = write_bdd_penalites(wb, funds)
    _, selected_count, fund_rows_count = build_exit_sheet(wb, src_ws, ws_cal.max_row, ws_pen.max_row, funds_by_isin)

    wb.active = wb.sheetnames.index("Calendrier de sortie")
    wb.save(out_path)
    print(f"OK -> {out_path} ({selected_count} fonds retenus sur {fund_rows_count} lignes analysées)")


if __name__ == "__main__":
    main()
