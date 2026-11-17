"""
Excel jadvalini TUMAN bo'yicha alohida fayllarga bo'ladi. Har bir tuman uchun
bitta .xlsx fayl: sarlavha (va ustidagi nom/title) saqlanadi, o'sha tuman
qatorlari ko'chiriladi, OXIRIGA talabalar soni yoziladi.

Tuman ustuni va sarlavha qatorini O'ZI topadi, shuning uchun turli tuzilishdagi
fayllarga ishlaydi (har biriga alohida ishga tushiring). Shrift, rang, chegara,
ustun kengligi va birlashgan kataklar saqlanadi.

O'rnatish:  pip install openpyxl
Ishlatish:  python tuman_boyicha_bolish.py [fayl.xlsx]
"""
import os
import re
import sys
import traceback
import unicodedata
from copy import copy
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = "1722 Дуал таълим талабалари.xlsx"   # yoki aniq fayl nomi
OUTPUT_DIR = None                  # None -> "<fayl nomi>_tumanlar" papkasi
COUNT_LABEL = "Jami talabalar soni:"
MAKE_INDEX = True                  # umumiy "00_Jami.xlsx" ro'yxat fayli
# ======================================================

if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]

TUMAN_HINT = re.compile(r"tuman|туман|shahar|шаҳар|rayon|район", re.I)
NOT_TUMAN = re.compile(r"viloyat|вилоят|oblast|область|mahalla|маҳалла|manzil", re.I)
HEAD_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def copy_style(src, dst):
    """Stilni boshqa faylga to'g'ri ko'chiradi (obyektlar bo'yicha)."""
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def norm_key(v):
    """Tuman nomini guruhlash kaliti (bo'shliq/registr farqi yo'q)."""
    return " ".join(nfc(v).casefold().split())


def safe_filename(s):
    s = nfc(s).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s or "tuman_yoq"


def detect_sheet_and_header(wb):
    """Tuman ustuni bor varaq, sarlavha qatori va tuman ustunini topadi."""
    order = [wb.active] + [w for w in wb.worksheets if w is not wb.active]
    best = None
    for ws in order:
        for r in range(1, min(ws.max_row, 8) + 1):
            cells = [c for c in ws[r] if c.value not in (None, "")]
            if len(cells) < 3:
                continue
            tcol = None
            for c in cells:
                t = nfc(c.value)
                if TUMAN_HINT.search(t) and not NOT_TUMAN.search(t):
                    tcol = c.column
                    break
            if tcol:
                score = len(cells) * 1000 + (ws.max_row)
                if best is None or score > best[0]:
                    best = (score, ws, r, tcol)
                break
    if best:
        return best[1], best[2], best[3]
    return None, None, None


def main():
    path = INPUT_FILE if os.path.isabs(INPUT_FILE) else os.path.join(BASE_DIR, INPUT_FILE)
    if not os.path.exists(path):
        print("Fayl topilmadi:", path, "\nPapkadagi Excel fayllar:")
        for f in sorted(os.listdir(BASE_DIR)):
            if f.lower().endswith((".xlsx", ".xlsm")):
                print("  ", f)
        return
    print("Fayl:", os.path.basename(path))
    wb = load_workbook(path)
    ws, hdr, tcol = detect_sheet_and_header(wb)
    if not ws:
        print("Tuman ustuni topilmadi.")
        return
    ncol = max((c.column for c in ws[hdr] if c.value not in (None, "")),
               default=ws.max_column)
    print(f"Varaq: [{ws.title}] | sarlavha qatori: {hdr} | "
          f"tuman ustuni: {get_column_letter(tcol)} ({nfc(ws.cell(hdr, tcol).value)})")

    # tumanlar bo'yicha qator indekslari (tartib saqlanadi)
    groups = {}          # key -> {"name":..., "rows":[...]}
    for r in range(hdr + 1, ws.max_row + 1):
        # butun qator bo'shmi?
        if all(ws.cell(r, c).value in (None, "") for c in range(1, ncol + 1)):
            continue
        tv = ws.cell(r, tcol).value
        key = norm_key(tv)
        if not key:
            key = "__bosh__"
            disp = "Tuman ko'rsatilmagan"
        else:
            disp = nfc(tv).strip()
        g = groups.setdefault(key, {"name": disp, "rows": []})
        g["rows"].append(r)

    if not groups:
        print("Ma'lumot qatori topilmadi.")
        return

    out_dir = OUTPUT_DIR or os.path.join(
        BASE_DIR, os.path.splitext(os.path.basename(path))[0] + "_tumanlar")
    os.makedirs(out_dir, exist_ok=True)

    # ustidagi title/header qatorlarining birlashgan kataklari
    top_merges = [mr for mr in ws.merged_cells.ranges if mr.max_row <= hdr]

    summary = []
    for key, g in sorted(groups.items(), key=lambda kv: kv[1]["name"]):
        nb = Workbook()
        ns = nb.active
        ns.title = safe_filename(g["name"])[:31] or "Tuman"

        # 1) title + sarlavha qatorlari (1..hdr) - stil bilan
        for r in range(1, hdr + 1):
            for c in range(1, ncol + 1):
                src = ws.cell(r, c)
                dst = ns.cell(r, c)
                dst.value = src.value
                copy_style(src, dst)
            if ws.row_dimensions[r].height:
                ns.row_dimensions[r].height = ws.row_dimensions[r].height
        for mr in top_merges:
            if mr.max_col <= ncol:
                ns.merge_cells(str(mr))

        # 2) ma'lumot qatorlari
        out_r = hdr + 1
        for r in g["rows"]:
            for c in range(1, ncol + 1):
                src = ws.cell(r, c)
                dst = ns.cell(out_r, c)
                dst.value = src.value
                copy_style(src, dst)
            if ws.row_dimensions[r].height:
                ns.row_dimensions[out_r].height = ws.row_dimensions[r].height
            out_r += 1

        # 3) oxiriga talabalar soni
        cnt = len(g["rows"])
        ns.cell(out_r, 1).value = f"{COUNT_LABEL} {cnt}"
        last_col = get_column_letter(ncol)
        ns.merge_cells(f"A{out_r}:{last_col}{out_r}")
        cell = ns.cell(out_r, 1)
        cell.font = Font(name="Times New Roman", size=12, bold=True)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.fill = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
        for c in range(1, ncol + 1):
            ns.cell(out_r, c).border = BOX

        # ustun kengligi + filtr + muzlatish
        for c in range(1, ncol + 1):
            L = get_column_letter(c)
            if ws.column_dimensions[L].width:
                ns.column_dimensions[L].width = ws.column_dimensions[L].width
        ns.auto_filter.ref = f"A{hdr}:{last_col}{out_r - 1}"
        ns.freeze_panes = f"A{hdr + 1}"
        ns.print_title_rows = f"1:{hdr}"
        ns.page_setup.orientation = "landscape"
        ns.page_setup.fitToWidth = 1
        ns.page_setup.fitToHeight = 0
        ns.sheet_properties.pageSetUpPr.fitToPage = True

        fname = f"{safe_filename(g['name'])}.xlsx"
        nb.save(os.path.join(out_dir, fname))
        summary.append((g["name"], cnt))
        print(f"  {g['name']:28} -> {cnt} ta  ({fname})")

    total = sum(c for _, c in summary)
    print(f"Jami: {len(summary)} ta tuman, {total} ta talaba")
    print("Papka:", out_dir)

    if MAKE_INDEX:
        ib = Workbook()
        iws = ib.active
        iws.title = "Jami"
        iws.append(["Tuman", "Talabalar soni"])
        for c in iws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BOX
        for name, cnt in sorted(summary, key=lambda x: -x[1]):
            iws.append([name, cnt])
            iws.cell(iws.max_row, 1).border = BOX
            iws.cell(iws.max_row, 2).border = BOX
        iws.append(["JAMI", total])
        for c in iws[iws.max_row]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
            c.border = BOX
        iws.column_dimensions["A"].width = 32
        iws.column_dimensions["B"].width = 16
        iws.freeze_panes = "A2"
        iws.auto_filter.ref = f"A1:B{iws.max_row}"
        ib.save(os.path.join(out_dir, "00_Jami_royxat.xlsx"))
        print("Umumiy ro'yxat: 00_Jami_royxat.xlsx")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
