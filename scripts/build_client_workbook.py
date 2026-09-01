#!/usr/bin/env python3
"""
Ajoute au classeur de consolidation client (ex : source/ConsolidationTemplateAlthosAI_V5.xlsx)
une nouvelle feuille "Calendrier de sortie" — un tableau COMPACT listant uniquement les fonds
que le client détient réellement (montant non nul dans sa consolidation) ET pour lesquels un
calendrier de RACHAT (sortie) officiel est connu dans la base Althos. Ni les fonds cotés, ni les
fonds non détenus, ni les fonds sans calendrier de rachat n'apparaissent : ce n'est pas une copie
de la feuille "Consolidation", juste la liste utile pour répondre à un client qui demande ses
délais de sortie. Pour chaque fonds retenu :

  - Titulaire (Monsieur / Madame / société... — vide si le fichier n'a qu'un seul titulaire)
  - Fonds, ISIN
  - Date d'investissement (à saisir par le conseiller)
  - Rachat — ordre avant / VL / exécuté / publié / cash reçu (valeurs reprises telles quelles
    de la base, pour la prochaine échéance de rachat à partir d'aujourd'hui)
  - Pénalité de sortie (vide si le délai de pénalité est dépassé ; message dans tous les autres
    cas, calculé par rapport à la date d'investissement saisie)

Présentation reprise de la feuille Consolidation elle-même (même police, mêmes couleurs) :
bandeaux de catégorie (beige, gras) au-dessus des fonds qu'ils regroupent, en-tête de tableau
dans la même couleur que celui de Consolidation. Un même fonds détenu par plusieurs titulaires
(ex. Monsieur ET Madame, chacun via son propre contrat) donne une ligne par titulaire.

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
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_data import (  # noqa: E402
    load_suivi, merge_extra_from_calendriers_par_fonds, load_calendrier,
    load_calendrier_par_fonds, merge_calendriers,
)

CAL_FILE = ROOT / "source" / "Calendriers_de_fonds_Althos.xlsx"



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
    """Une ligne de catégorie (bandeau) est en gras + fond uni, sans ISIN en colonne B.

    Associe aussi à chaque ligne fonds le libellé de la dernière catégorie rencontrée
    au-dessus d'elle (row_to_category), pour reproduire les mêmes bandeaux de catégorie
    dans la feuille générée.
    """
    cat_rows, fund_rows = [], []
    row_to_category = {}
    current_category = None
    for r in range(header_row + 2, ws.max_row + 1):  # header_row+1 = ligne de total général
        a = ws.cell(row=r, column=1)
        b = ws.cell(row=r, column=2)
        if a.value is None:
            continue
        if a.fill.patternType == "solid" and (a.font.bold or False) and b.value is None:
            cat_rows.append(r)
            current_category = str(a.value).strip()
        else:
            fund_rows.append(r)
            row_to_category[r] = current_category
    return cat_rows, fund_rows, row_to_category


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


def detect_owner_labels(ws, header_row, total_col):
    """Libellé du titulaire (Monsieur / Madame / société / prénom...) par colonne 'Contrat',
    lu sur la ligne juste au-dessus de l'en-tête (souvent fusionnée sur plusieurs colonnes
    contrat). openpyxl ne reporte la valeur que sur la cellule ancre d'une fusion : on résout
    donc explicitement les plages fusionnées qui couvrent cette ligne. Absent ou vide pour un
    client à titulaire unique (pas de subdivision) : toutes les colonnes retombent alors sur "".
    """
    labels = {}
    owner_row = header_row - 1
    if not total_col or owner_row < 1:
        return labels
    merged_lookup = {}
    for mc in ws.merged_cells.ranges:
        if mc.min_row <= owner_row <= mc.max_row:
            anchor_val = ws.cell(row=mc.min_row, column=mc.min_col).value
            for c in range(mc.min_col, mc.max_col + 1):
                merged_lookup[c] = anchor_val
    for c in range(3, total_col):
        v = merged_lookup.get(c, ws.cell(row=owner_row, column=c).value)
        labels[c] = str(v).strip() if v is not None else ""
    return labels


def holding_by_owner(ws, row, total_col, owner_labels):
    """Répartit le montant détenu d'une ligne fonds par titulaire (clé "" si non subdivisé)."""
    if not total_col:
        return [("", None)]  # structure inconnue : une seule ligne, montant indéterminé
    amounts = {}
    for c in range(3, total_col):
        v = ws.cell(row=row, column=c).value
        if isinstance(v, (int, float)):
            label = owner_labels.get(c, "")
            amounts[label] = amounts.get(label, 0) + v
    return list(amounts.items())


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

def format_date_fr(d):
    """« 18/08/2026 », au format numérique jour/mois/année."""
    return d.strftime("%d/%m/%Y")


HELPER_NAMES = [
    "has_sortie", "next_cutoff_sortie", "next_val_sortie", "next_exec_sortie",
    "next_pub_sortie", "next_cash_sortie",
    "months_held", "pen_found", "kind", "raw",
    "max1", "rate1", "max2", "rate2", "max3", "rate3", "max4", "rate4", "rate5",
    "rate_now",
]
HELPER_FIRST_COL = 11  # K (colonnes visibles jusqu'en J désormais : Titulaire ajouté en colonne A)


def helper_col(name):
    idx = HELPER_FIRST_COL + HELPER_NAMES.index(name)
    return get_column_letter(idx)


def has_rachat_calendar(calendar, isin):
    """Un fonds n'a de calendrier "de sortie" que s'il a des échéances de type Rachat."""
    by_type = calendar.get(isin) if isin else None
    return bool(by_type and (by_type.get("Rachat") or by_type.get("Souscription et rachat")))


def build_exit_sheet(wb, src_ws, calendar, cal_last_row, pen_last_row, funds_by_isin):
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
    cat_rows, fund_rows, row_to_category = classify_rows(src_ws, header_row)
    owner_labels = detect_owner_labels(src_ws, header_row, total_col)

    # Fonds détenus (répartis par titulaire de contrat) ET dotés d'un calendrier de RACHAT connu.
    # Un même fonds détenu par plusieurs titulaires (ex. Monsieur ET Madame, chacun via son
    # propre contrat) donne une ligne par titulaire, pour une date d'investissement et un statut
    # de pénalité propres à chacun.
    selected = []
    for r in fund_rows:
        isin_raw = src_ws.cell(row=r, column=2).value
        isin = isin_raw.strip() if isinstance(isin_raw, str) else isin_raw
        fund = funds_by_isin.get(isin) if isin else None
        # Un fonds fermé (aucune sortie possible hors cas exceptionnel) est toujours retenu même
        # sans calendrier de rachat : c'est justement l'information à signaler au conseiller.
        is_ferme = bool(fund and fund.get("penalite", {}).get("kind") == "ferme")
        if not fund or not (is_ferme or has_rachat_calendar(calendar, isin)):
            continue
        nom = src_ws.cell(row=r, column=1).value or fund["nom"]
        category = row_to_category.get(r) or ""
        for owner, amount in holding_by_owner(src_ws, r, total_col, owner_labels):
            if amount is not None and abs(amount) <= 0.005:
                continue
            selected.append({"isin": isin, "nom": nom, "category": category, "owner": owner})

    # Présentation "à la Althos" : bandeau de titre + sous-titre daté (repris tel quel des lignes
    # 1 et 3 de Consolidation — même couleur, même police, même mise en italique), en-tête de
    # tableau dans le même bleu que Consolidation et le titre, bandeaux de catégorie en beige
    # (repris de la couleur des catégories dans Consolidation), fonds sur fond blanc, quadrillage
    # fin en beige. Police commune à toute la feuille (celle de l'en-tête Consolidation, pas un
    # choix arbitraire).
    header_style = copy(src_ws.cell(row=header_row, column=1)._style)  # navy header (theme1)
    header_font = src_ws.cell(row=header_row, column=1).font
    data_font = Font(name=header_font.name, size=header_font.size, bold=False)

    title_style = copy(src_ws.cell(row=1, column=1)._style)
    subtitle_style = copy(src_ws.cell(row=3, column=1)._style)
    # Style de bandeau de catégorie : repris tel quel de la première ligne de catégorie trouvée
    # dans la Consolidation (même couleur beige, même police en gras).
    category_style = copy(src_ws.cell(row=cat_rows[0], column=1)._style) if cat_rows else None
    # Couleur de quadrillage : reprise telle quelle (référence de thème + teinte, PAS une valeur
    # fixe) de la bordure déjà utilisée par l'en-tête de Consolidation — sa couleur réelle dépend
    # du thème propre à chaque classeur (beige dans certains, bleu dans d'autres), donc on ne peut
    # pas la coder en dur.
    header_border = src_ws.cell(row=header_row, column=1).border
    ref_side = header_border.left or header_border.top or header_border.right or header_border.bottom
    grid_side = (copy(ref_side) if ref_side is not None and ref_side.style
                 else Side(style="thin", color="FFDDCCB8"))
    grid_side.style = "medium"
    grid_border = Border(top=grid_side, left=grid_side, bottom=grid_side, right=grid_side)
    white_fill = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")
    ws.sheet_view.showGridLines = False  # comme Consolidation : pas de quadrillage Excel par défaut
    # Vue "aperçu des sauts de page", comme Consolidation : c'est ce qui affiche automatiquement
    # le contour bleu de la zone d'impression en vue Normale.
    ws.sheet_view.view = "pageBreakPreview"

    headers = ["Titulaire", "Fonds", "ISIN", "Date d'investissement",
               "Rachat — ordre avant", "Rachat — VL", "Rachat — exécuté",
               "Rachat — publié", "Rachat — cash reçu", "Pénalité de sortie"]
    HEADER_ROW_OUT = 5  # 1: titre, 2: (vide), 3: sous-titre daté, 4: (vide), 5: en-tête
    FIRST_DATA_ROW = HEADER_ROW_OUT + 1

    title_cell = ws.cell(row=1, column=1, value="Calendrier des délais de sortie")
    title_cell._style = copy(title_style)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.row_dimensions[1].height = src_ws.row_dimensions[1].height or 22.5

    subtitle_cell = ws.cell(row=3, column=1, value=format_date_fr(datetime.date.today()))
    subtitle_cell._style = copy(subtitle_style)
    subtitle_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(headers))
    ws.row_dimensions[3].height = src_ws.row_dimensions[3].height or 14

    for i, label in enumerate(headers):
        cell = ws.cell(row=HEADER_ROW_OUT, column=i + 1, value=label)
        cell._style = copy(header_style)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.border = grid_border
    ws.row_dimensions[HEADER_ROW_OUT].height = 22

    widths = [14, 32, 14, 22, 22, 15, 19, 18, 20, 46]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    for name in HELPER_NAMES:
        ws.column_dimensions[helper_col(name)].hidden = True

    if not selected:
        msg = ("Aucun fonds avec calendrier de sortie (rachat) connu n'est actuellement détenu par "
               'ce client (montant total nul, ou fonds hors périmètre "fonds non cotés suivis").')
        ws.cell(row=FIRST_DATA_ROW, column=1, value=msg)
        ws.merge_cells(start_row=FIRST_DATA_ROW, start_column=1, end_row=FIRST_DATA_ROW, end_column=len(headers))
        ws.cell(row=FIRST_DATA_ROW, column=1).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[FIRST_DATA_ROW].height = 30
        apply_print_setup(ws, FIRST_DATA_ROW, len(headers))
        return ws, 0, len(fund_rows)

    r = FIRST_DATA_ROW
    current_category = object()  # sentinelle : force le 1er bandeau même si catégorie ""
    for item in selected:
        if category_style is not None and item["category"] != current_category:
            current_category = item["category"]
            cat_cell = ws.cell(row=r, column=1, value=current_category or "Autres fonds")
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c)._style = copy(category_style)
                ws.cell(row=r, column=c).border = grid_border
            cat_cell.alignment = Alignment(vertical="center")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers))
            ws.row_dimensions[r].height = 20
            r += 1

        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).fill = white_fill
            ws.cell(row=r, column=c).border = grid_border

        isin, nom, owner = item["isin"], item["nom"], item["owner"]
        ws.cell(row=r, column=1, value=owner).font = data_font
        ws.cell(row=r, column=2, value=nom).font = data_font
        ws.cell(row=r, column=3, value=isin).font = data_font
        for col in (1, 2, 3):
            ws.cell(row=r, column=col).alignment = Alignment(vertical="center")

        date_cell = ws.cell(row=r, column=4)
        date_cell.number_format = "dd/mm/yyyy"
        date_cell.font = data_font

        b = f'"{isin}"'
        c = f"$D{r}"

        L, M, N = helper_col("has_sortie"), helper_col("next_cutoff_sortie"), helper_col("next_val_sortie")
        Nx, Np, O = helper_col("next_exec_sortie"), helper_col("next_pub_sortie"), helper_col("next_cash_sortie")
        Q = helper_col("months_held")
        R = helper_col("pen_found")
        S = helper_col("kind")
        Tc = helper_col("raw")
        U1, V1, W1, X1, Y1, Z1, AA1, AB1, AC1 = (helper_col("max1"), helper_col("rate1"), helper_col("max2"),
                                                  helper_col("rate2"), helper_col("max3"), helper_col("rate3"),
                                                  helper_col("max4"), helper_col("rate4"), helper_col("rate5"))
        AD1 = helper_col("rate_now")

        ws[f"{L}{r}"] = f'=COUNTIFS({cal("A")},{b},{cal("C")},"Rachat")'
        ws[f"{M}{r}"] = (f'=IF({L}{r}=0,"",IFERROR(_xlfn.MINIFS({cal("D")},'
                          f'{cal("A")},{b},{cal("C")},"Rachat",'
                          f'{cal("D")},">="&TODAY()),""))')
        # VL / exécuté / publié / cash reçu de CETTE échéance précise : on réutilise MINIFS avec
        # une égalité exacte sur la date de cut-off déjà trouvée (au lieu d'une reconstruction de
        # clé texte + MATCH, plus fragile) — même mécanisme que {M} ci-dessus, qui fonctionne de
        # façon fiable. MINIFS ignore les cellules vides : si ce champ précis n'est pas renseigné
        # dans la base pour cette échéance, MINIFS ne trouve aucune valeur numérique et renvoie 0
        # (jamais une vraie erreur) — sans la vérification "=0" ci-dessous, Excel afficherait ce 0
        # comme une date, "00/01/1900", au lieu de laisser la cellule vide.
        def minifs_field(col):
            expr = f'_xlfn.MINIFS({cal(col)},{cal("A")},{b},{cal("C")},"Rachat",{cal("D")},{M}{r})'
            return f'=IF({M}{r}="","",IFERROR(IF({expr}=0,"",{expr}),""))'

        ws[f"{N}{r}"] = minifs_field("E")
        ws[f"{Nx}{r}"] = minifs_field("F")
        ws[f"{Np}{r}"] = minifs_field("G")
        ws[f"{O}{r}"] = minifs_field("H")

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
                            f'{Q}{r}<{Y1}{r},{Z1}{r},{Q}{r}<{AA1}{r},{AB1}{r},TRUE(),{AC1}{r}))')

        # Colonnes visibles E..I : valeurs reprises telles quelles de la base (une par champ).
        for col_letter, helper in (("E", M), ("F", N), ("G", Nx), ("H", Np), ("I", O)):
            cell = ws[f"{col_letter}{r}"]
            cell.value = f"={helper}{r}"
            cell.number_format = "dd/mm/yyyy"
            cell.font = data_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Pénalité de sortie : vide dès qu'il n'y a rien d'actionnable à signaler — pas de
        # pénalité prévue pour ce fonds (kind="aucune"), aucune pénalité renseignée dans la base
        # (kind="inconnue", traité comme "pas de pénalité"), ou délai de pénalité dépassé. Un
        # message n'apparaît que dans les cas où le conseiller doit agir : fonds fermé (sortie
        # impossible hors cas exceptionnel), règle ambiguë à vérifier à la main, ou pénalité
        # active à signaler au client.
        ws[f"J{r}"] = (
            f'=_xlfn.IFS('
            f'{S}{r}="ferme","🔒 FONDS FERMÉ — sortie non disponible (sauf cas exceptionnel prévu au règlement, ex. décès, invalidité). Vérifier le DICI du fonds.",'
            f'{S}{r}="manuel","⚠️ À VÉRIFIER MANUELLEMENT : "&{Tc}{r},'
            f'{S}{r}="aucune","",'
            f'{S}{r}="inconnue","",'
            f'{c}="","Saisir une date d\'investissement pour statuer sur la pénalité.",'
            f'{Q}{r}="FUTUR","Date d\'investissement postérieure à aujourd\'hui — vérifier la saisie.",'
            f'{AD1}{r}>0,"⚠️ CONCERNÉ : pénalité de "&{AD1}{r}&"% (détention "&{Q}{r}&" mois). "&{Tc}{r},'
            f'TRUE(),"")'
        )
        pen_cell = ws[f"J{r}"]
        pen_cell.font = data_font
        pen_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws.row_dimensions[r].height = 45
        r += 1

    apply_print_setup(ws, r - 1, len(headers))
    return ws, len(selected), len(fund_rows)


def apply_print_setup(ws, last_row, last_visible_col):
    """Zone d'impression = tout le tableau, mise à l'échelle sur une page en largeur, paysage —
    comme sur la feuille Consolidation. Marge de 1 ligne/colonne entre le contenu et le contour
    de la zone d'impression. Pour les colonnes, "le contenu" inclut les colonnes d'aide masquées
    (jusqu'à HELPER_FIRST_COL + len(HELPER_NAMES) - 1), pas seulement les colonnes visibles :
    sinon le contour passerait au milieu de colonnes masquées au lieu de se terminer 1 colonne
    après elles."""
    last_helper_col = HELPER_FIRST_COL + len(HELPER_NAMES) - 1
    last_col = max(last_visible_col, last_helper_col)
    ws.print_area = f"A1:{get_column_letter(last_col + 1)}{last_row + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


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
    # Force Excel à recalculer TOUTES les formules à l'ouverture (le fichier n'a pas de
    # calcChain.xml puisqu'openpyxl n'exécute pas les formules lui-même) : évite les cellules
    # calculées qui restent vides tant que l'utilisateur n'a pas forcé un recalcul manuel.
    wb.calculation.fullCalcOnLoad = True

    ws_cal = write_bdd_calendrier(wb, calendar)
    ws_pen = write_bdd_penalites(wb, funds)
    _, selected_count, fund_rows_count = build_exit_sheet(wb, src_ws, calendar, ws_cal.max_row, ws_pen.max_row, funds_by_isin)

    wb.active = wb.sheetnames.index("Calendrier de sortie")
    wb.save(out_path)
    print(f"OK -> {out_path} ({selected_count} fonds retenus sur {fund_rows_count} lignes analysées)")


if __name__ == "__main__":
    main()
