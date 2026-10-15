# -*- coding: utf-8 -*-
"""
Export Excel des inscriptions aux ateliers de robotique
المديرية الإقليمية نواصر — الثانوية التأهيلية علال بن عبد الله
=================================================================
Génère un classeur .xlsx en arabe (feuilles RTL) contenant :
  - Feuille 1 « التسجيلات »  : liste complète des inscrits (id, nom, prénom,
    classe, فوج, jour1, jour2, créneau, date) avec ligne de total et en-têtes figés.
  - Feuille 2 « إحصائيات » : répartition par فوج (places restantes, taux de
    remplissage), par قسم (classe), par حصة (jour + créneau, avec capacité,
    places restantes et taux de remplissage) et par paire de jours (اليومان).

Style dérivé d'un design system « borderless-first » (3 couleurs max) :
PRIMARY teal de l'application, SECONDARY version claire, neutres chauds,
lignes alternées, pas de grille pleine.
"""

from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------
# Design tokens (alignés sur la charte de l'application)
# ------------------------------------------------------------
PRIMARY = "0F4C5C"      # teal profond : en-têtes, titres
SECONDARY = "D6E8ED"    # teal clair : ligne de total
NEUTRAL_900 = "37352F"  # texte courant
NEUTRAL_600 = "8C8A84"  # légendes
NEUTRAL_200 = "E9E9E8"  # filets (bas d'en-tête / haut de total)
NEUTRAL_100 = "F7F7F5"  # lignes impaires
NEUTRAL_0 = "FFFFFF"    # lignes paires

# Tahoma : excellent rendu de l'arabe sous Windows (Excel), macOS et LibreOffice
FONT_NAME = "Tahoma"

HEADERS = ["رقم", "الاسم العائلي", "الاسم الشخصي", "القسم", "الفوج", "اليوم الأول", "اليوم الثاني", "التوقيت", "تاريخ التسجيل"]
COL_WIDTHS = [8, 18, 18, 12, 16, 12, 12, 14, 21]

GROUPS = ["الفوج الأول", "الفوج الثاني", "الفوج الثالث", "الفوج الرابع"]
CLASSES = [
    "TCLSH-1",
    "TCSF-1", "TCSF-2", "TCSF-3", "TCSF-4",
    "TCSF-5", "TCSF-6", "TCSF-7", "TCSF-8", "TCSF-9", "TCSF-10",
]

# Jours et créneaux des ateliers (identiques à app.py) — séances d'une heure
DAYS = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]
TIME_SLOTS = [
    "14:00 – 15:00",
    "15:00 – 16:00",
    "16:00 – 17:00",
    "17:00 – 18:00",
]

# Paires de jours possibles pour les 2 séances hebdomadaires (identiques à app.py)
DAY_PAIRS = [
    ("الاثنين", "الثلاثاء"),
    ("الاثنين", "الأربعاء"),
    ("الاثنين", "الخميس"),
    ("الثلاثاء", "الأربعاء"),
    ("الثلاثاء", "الخميس"),
    ("الأربعاء", "الخميس"),
]


def _font(size=11, bold=False, color=NEUTRAL_900):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color)


def _setup_sheet(ws, title, subtitle, last_col):
    """Mise en page commune : origine B2, titre, sous-titre, sans grille."""
    ws.sheet_view.rightToLeft = True  # feuille en droite-à-gauche (arabe)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3  # marge visuelle
    ws.row_dimensions[1].height = 15

    # Titre (B2, fusionné)
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=last_col)
    c = ws.cell(row=2, column=2, value=title)
    c.font = _font(16, bold=True, color=PRIMARY)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 32

    # Sous-titre (B3, fusionné)
    ws.merge_cells(start_row=3, start_column=2, end_row=3, end_column=last_col)
    s = ws.cell(row=3, column=2, value=subtitle)
    s.font = _font(9, color=NEUTRAL_600)
    s.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[3].height = 14


def _style_header(ws, row_num, headers, start_col=2):
    """Ligne d'en-têtes : fond PRIMARY, texte blanc, filet bas."""
    header_border = Border(bottom=Side(style="thin", color=NEUTRAL_200))
    for i, h in enumerate(headers):
        cell = ws.cell(row=row_num, column=start_col + i, value=h)
        cell.fill = PatternFill("solid", fgColor=PRIMARY)
        cell.font = _font(11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border
    ws.row_dimensions[row_num].height = 28


def _style_data_cell(cell, row_index, horizontal="right"):
    """Cellule de données : fond alterné, sans bordures."""
    fill = NEUTRAL_0 if row_index % 2 == 0 else NEUTRAL_100
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.font = _font(11)
    cell.alignment = Alignment(horizontal=horizontal, vertical="center")


def _style_total_row(ws, row_num, start_col, end_col):
    """Ligne de total : fond SECONDARY, filet haut medium."""
    top_border = Border(top=Side(style="medium", color=NEUTRAL_200))
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = PatternFill("solid", fgColor=SECONDARY)
        cell.font = _font(11, bold=True, color=PRIMARY)
        cell.border = top_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_num].height = 26


def build_excel_bytes(rows, max_per_group=20, max_per_slot=20):
    """Construit le classeur Excel (bytes) à partir des inscriptions.

    Args:
        rows: liste de dictionnaires {id, nom, prenom, classe, groupe, jour1,
              jour2, horaire, date_inscription} (compat : ancien champ jour)
        max_per_group: capacité maximale par فوج (pour les statistiques)
        max_per_slot: capacité maximale par حصة (jour + créneau)

    Returns:
        bytes du fichier .xlsx
    """
    wb = Workbook()

    # ==========================================================
    # FEUILLE 1 — قائمة التسجيلات (liste des inscrits)
    # ==========================================================
    ws = wb.active
    ws.title = "التسجيلات"
    last_col = 1 + len(HEADERS)  # B..J
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    subtitle = (
        "الثانوية التأهيلية علال بن عبد الله — المديرية الإقليمية نواصر"
        f" | عدد التسجيلات: {len(rows)} | استخراج بتاريخ: {stamp}"
    )
    _setup_sheet(ws, "قائمة التسجيلات في ورشات التكوين في الروبوتيك", subtitle, last_col)

    # Largeurs de colonnes
    for i, w in enumerate(COL_WIDTHS):
        ws.column_dimensions[get_column_letter(2 + i)].width = w

    # En-têtes (ligne 4)
    _style_header(ws, 4, HEADERS)

    # Données (à partir de la ligne 5)
    for r_idx, row in enumerate(rows):
        row_num = 5 + r_idx
        values = [
            row.get("id"),
            row.get("nom", ""),
            row.get("prenom", ""),
            row.get("classe", ""),
            row.get("groupe", ""),
            row.get("jour1", row.get("jour", "")),
            row.get("jour2", ""),
            row.get("horaire", ""),
            row.get("date_inscription", ""),
        ]
        for c_idx, v in enumerate(values):
            cell = ws.cell(row=row_num, column=2 + c_idx, value=v)
            # centrer : id, classe, jours, créneau, date — droite : textes arabes
            horizontal = "center" if c_idx in (0, 3, 5, 6, 7, 8) else "right"
            _style_data_cell(cell, r_idx, horizontal=horizontal)
        ws.row_dimensions[row_num].height = 22

    # Ligne de total
    total_row = 5 + len(rows)
    ws.cell(row=total_row, column=2, value="المجموع")
    ws.cell(row=total_row, column=3, value=len(rows))
    _style_total_row(ws, total_row, 2, last_col)

    # En-têtes figés au défilement + mise en page d'impression
    ws.freeze_panes = "A5"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "4:4"

    # ==========================================================
    # FEUILLE 2 — إحصائيات (statistiques par فوج et par قسم)
    # ==========================================================
    ws2 = wb.create_sheet("إحصائيات")
    _setup_sheet(ws2, "إحصائيات التسجيلات", subtitle, 7)

    counts_group = {g: 0 for g in GROUPS}
    counts_class = {c: 0 for c in CLASSES}
    for row in rows:
        g = row.get("groupe")
        c = row.get("classe")
        if g in counts_group:
            counts_group[g] += 1
        if c in counts_class:
            counts_class[c] += 1

    # ---- Section A : répartition par فوج ----
    ws2.cell(row=5, column=2, value="التوزيع حسب الفوج").font = _font(12, bold=True, color=PRIMARY)
    ws2.row_dimensions[5].height = 24

    headers_a = ["الفوج", "عدد المسجلين", "المقاعد المتبقية", "نسبة الامتلاء"]
    _style_header(ws2, 6, headers_a)
    for i, w in enumerate([16, 15, 16, 14, 17, 14]):  # B..G (commun aux 3 sections)
        ws2.column_dimensions[get_column_letter(2 + i)].width = w

    r = 7
    for idx, g in enumerate(GROUPS):
        n = counts_group[g]
        remain = max_per_group - n
        ratio = n / max_per_group if max_per_group else 0
        vals = [g, n, remain, ratio]
        for c_idx, v in enumerate(vals):
            cell = ws2.cell(row=r, column=2 + c_idx, value=v)
            _style_data_cell(cell, idx, horizontal="center")
            if c_idx == 3:
                cell.number_format = "0%"
        ws2.row_dimensions[r].height = 22
        r += 1

    ws2.cell(row=r, column=2, value="المجموع")
    ws2.cell(row=r, column=3, value=sum(counts_group.values()))
    _style_total_row(ws2, r, 2, 5)

    # ---- Section B : répartition par قسم ----
    sec_b = r + 3
    ws2.cell(row=sec_b, column=2, value="التوزيع حسب القسم").font = _font(12, bold=True, color=PRIMARY)
    ws2.row_dimensions[sec_b].height = 24

    headers_b = ["القسم", "عدد المسجلين"]
    _style_header(ws2, sec_b + 1, headers_b)

    r = sec_b + 2
    for idx, c in enumerate(CLASSES):
        cell1 = ws2.cell(row=r, column=2, value=c)
        _style_data_cell(cell1, idx, horizontal="center")
        cell2 = ws2.cell(row=r, column=3, value=counts_class[c])
        _style_data_cell(cell2, idx, horizontal="center")
        ws2.row_dimensions[r].height = 22
        r += 1

    ws2.cell(row=r, column=2, value="المجموع")
    ws2.cell(row=r, column=3, value=sum(counts_class.values()))
    _style_total_row(ws2, r, 2, 3)

    # ---- Section C : répartition par حصة (jour + créneau) ----
    # Chaque inscription compte pour ses 2 séances (jour1 + jour2, même horaire)
    counts_slot = {}
    for row in rows:
        h = row.get("horaire", "")
        jours = [j for j in (row.get("jour1"), row.get("jour2")) if j]
        if not jours and row.get("jour"):
            jours = [row.get("jour")]  # compat ancien format
        for j in jours:
            if j in DAYS and h in TIME_SLOTS:
                counts_slot[(j, h)] = counts_slot.get((j, h), 0) + 1

    sec_c = r + 3
    ws2.cell(row=sec_c, column=2, value="التوزيع حسب الحصة (اليوم والتوقيت)").font = _font(
        12, bold=True, color=PRIMARY
    )
    ws2.row_dimensions[sec_c].height = 24

    headers_c = ["اليوم", "التوقيت", "عدد المسجلين", "السعة", "المقاعد المتبقية", "نسبة الامتلاء"]
    _style_header(ws2, sec_c + 1, headers_c)

    r = sec_c + 2
    idx = 0
    for j in DAYS:
        for h in TIME_SLOTS:
            n = counts_slot.get((j, h), 0)
            remain = max_per_slot - n
            ratio = n / max_per_slot if max_per_slot else 0
            for c_idx, v in enumerate([j, h, n, max_per_slot, remain, ratio]):
                cell = ws2.cell(row=r, column=2 + c_idx, value=v)
                _style_data_cell(cell, idx, horizontal="center")
                if c_idx == 5:
                    cell.number_format = "0%"
            ws2.row_dimensions[r].height = 22
            r += 1
            idx += 1

    slots_capacity = len(DAYS) * len(TIME_SLOTS) * max_per_slot
    n_slots_total = sum(counts_slot.values())
    ws2.cell(row=r, column=2, value="المجموع")
    ws2.cell(row=r, column=4, value=n_slots_total)
    ws2.cell(row=r, column=5, value=slots_capacity)
    ws2.cell(row=r, column=6, value=slots_capacity - n_slots_total)
    cell_ratio = ws2.cell(
        row=r, column=7, value=(n_slots_total / slots_capacity) if slots_capacity else 0
    )
    _style_total_row(ws2, r, 2, 7)
    cell_ratio.number_format = "0%"

    # ---- Section D : répartition par paire de jours (اليومان) ----
    counts_pair = {}
    for row in rows:
        j1 = row.get("jour1") or ""
        j2 = row.get("jour2") or ""
        if j1 in DAYS and j2 in DAYS:
            counts_pair[(j1, j2)] = counts_pair.get((j1, j2), 0) + 1

    sec_d = r + 3
    ws2.cell(row=sec_d, column=2, value="التوزيع حسب أيام الحصتين (اليومان)").font = _font(
        12, bold=True, color=PRIMARY
    )
    ws2.row_dimensions[sec_d].height = 24

    headers_d = ["اليوم الأول", "اليوم الثاني", "عدد المسجلين"]
    _style_header(ws2, sec_d + 1, headers_d)

    r = sec_d + 2
    idx = 0
    for d1, d2 in DAY_PAIRS:
        n = counts_pair.get((d1, d2), 0)
        if d1 != d2:
            n += counts_pair.get((d2, d1), 0)
        for c_idx, v in enumerate([d1, d2, n]):
            cell = ws2.cell(row=r, column=2 + c_idx, value=v)
            _style_data_cell(cell, idx, horizontal="center")
        ws2.row_dimensions[r].height = 22
        r += 1
        idx += 1

    ws2.cell(row=r, column=2, value="المجموع")
    ws2.cell(row=r, column=4, value=sum(counts_pair.values()))
    _style_total_row(ws2, r, 2, 4)

    # Notes de bas de tableau
    note = ws2.cell(
        row=r + 2,
        column=2,
        value=(
            f"الطاقة الاستيعابية: {len(GROUPS)} أفواج × {max_per_group} تلميذا = "
            f"{len(GROUPS) * max_per_group} — بحد أقصى {max_per_slot} تلاميذا في كل حصة"
        ),
    )
    note.font = _font(9, color=NEUTRAL_600)
    note.alignment = Alignment(horizontal="right", vertical="center")
    note2 = ws2.cell(
        row=r + 3,
        column=2,
        value="موعد الورشات: انطلاقا من الاثنين 26 أكتوبر 2026 — حصتان في الأسبوع يومين مختلفين في نفس التوقيت — من الاثنين إلى الخميس، بين 14:00 و18:00 — مدة كل حصة ساعة واحدة",
    )
    note2.font = _font(9, color=NEUTRAL_600)
    note2.alignment = Alignment(horizontal="right", vertical="center")

    wb.properties.creator = "Z.ai"

    from io import BytesIO

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def save_excel_file(rows, path, max_per_group=20, max_per_slot=20):
    """Sauvegarde locale du classeur (inscriptions.xlsx à côté de l'app)."""
    with open(path, "wb") as f:
        f.write(
            build_excel_bytes(rows, max_per_group=max_per_group, max_per_slot=max_per_slot)
        )
