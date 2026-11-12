# -*- coding: utf-8 -*-
"""
"Mahalla" ustunining yoniga "Mahalla nomi" ustunini qo'shadi va barcha yozuvlarni
bitta stilga keltiradi:   "SHARQGULI" MFY

Qaysi ko'rinishlarni to'g'irlaydi (misollar):
  m.f.y. Yangihayot              -> "YANGIHAYOT" MFY
  MFY "SHARQGULI"                -> "SHARQGULI" MFY
  Yangiobod MFY / Yangiobod mfy  -> "YANGIOBOD" MFY
  Yangiobod mahallasi            -> "YANGIOBOD" MFY
  Yangiobod mahalla fuqarolar yig'ini -> "YANGIOBOD" MFY
  «Bahor» MFY / ''Bahor'' MFY    -> "BAHOR" MFY
  "Navro'z" MFY / Navroʻz MFY    -> "NAVRO`Z" MFY   (tutuq belgisi bir xil qilinadi)
  Янгиобод МФЙ / м.ф.й. Янгиобод -> "ЯНГИОБОД" MFY

Qo'shimcha:
  - Skriptni qayta ishga tushirsangiz, ustun ikkinchi marta qo'shilmaydi (eskisi yangilanadi).
  - "Tekshiruv" varag'i yaratiladi: har bir mahalla necha marta uchragani va
    yozilishi juda o'xshash nomlar (masalan CHINOBOD / CHINABOD) — qo'lda tekshirish uchun.
  - Natija yangi faylga yoziladi (<nom>_mahalla.xlsx), asl fayl o'zgarmaydi.

Ishga tushirish:
  py -m pip install openpyxl
  py mahalla_nomi.py
"""

import os
import re
import sys
from copy import copy
from difflib import SequenceMatcher

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

# ================== ASOSIY SOZLAMALAR ==================
DEFAULT_INPUT = "Yatt o'tmagan 02.11.2026 18-00.xlsx"  # topilmasa, papkadagi fayllardan tanlatadi
SOURCE_HEADER = "mahalla"          # manba ustun sarlavhasi (katta-kichik harf farqi yo'q)
NEW_HEADER = "Mahalla nomi"        # yangi ustun sarlavhasi
APOSTROPHE = "`"                   # o`, g` uchun ishlatiladigan belgi: "`" yoki "'" yoki "ʻ"
SUFFIX = "MFY"                     # oxiriga qo'shiladigan qo'shimcha
UPPERCASE = True                   # nomni katta harflarda yozish
SIMILARITY = 0.85                  # "Tekshiruv"da o'xshash nomlarni ko'rsatish chegarasi (0..1)
# =======================================================

APOS_CHARS = "'`‘’ʻʼ´ʹ′"
QUOTE_CHARS = "\"«»“”„‟″"
PH = "\u0001"   # so'z ichidagi tutuq belgisi uchun vaqtinchalik belgi

# Olib tashlanadigan "mahalla" so'zlari (lotin va kirill). Uzunroqlari birinchi.
NOISE = [
    r"mahalla\s+fuqarolar(?:i)?\s+yig\W?ini(?:\s+hududida)?",
    r"mahalla\s+fuqarolar(?:i)?\s+yig\W?ilishi",
    r"fuqarolar\s+yig\W?ini",
    r"махалла\s+фуқаролар(?:и)?\s+йиғини", r"маҳалла\s+фуқаролар(?:и)?\s+йиғини",
    r"махалла\s+фукаролар(?:и)?\s+йигини", r"маҳалла\s+фукаролар(?:и)?\s+йигини",
    r"фуқаролар\s+йиғини", r"фукаролар\s+йигини",
    r"m\s*\.?\s*f\s*\.?\s*y\s*\.?(?:\s*i)?",           # m.f.y. / mfy / MFY / MFYi
    r"м\s*\.?\s*ф\s*\.?\s*й\s*\.?(?:\s*и)?",           # м.ф.й. / МФЙ
    r"mahallasi", r"mahalla", r"maxallasi", r"maxalla", r"mahalasi", r"mahala",
    r"маҳалласи", r"махалласи", r"маҳалла", r"махалла",
    r"hududida", r"hududi",
]
NOISE_RE = re.compile(r"(?<![\w" + PH + r"])(?:" + "|".join(NOISE) + r")(?![\w" + PH + r"])",
                      re.IGNORECASE)


def normalize(value):
    """Bitta katakdagi qiymatni '"NOM" MFY' ko'rinishiga keltiradi. Bo'sh bo'lsa ''."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s in ("0", "-", "—"):
        return ""

    s = s.replace(" ", " ")
    # 1) So'z ichidagi tutuq belgilarini saqlab qolamiz: o'z, g'alaba, Navro`z
    s = re.sub(r"(?<=\w)[" + re.escape(APOS_CHARS) + r"](?=\w)", PH, s)
    # O' / G' so'z oxirida bo'lsa ham (masalan "Ilg'or" emas, "...o'" holati kam) — o/g dan keyin
    s = re.sub(r"(?<=[oOgG])[" + re.escape(APOS_CHARS) + r"](?=\s|$)", PH, s)

    # 2) Qolgan barcha qo'shtirnoq/tirnoqlarni olib tashlaymiz
    s = re.sub("[" + re.escape(QUOTE_CHARS + APOS_CHARS) + "]", " ", s)

    # 3) "mahalla", "MFY", "m.f.y." kabi so'zlarni olib tashlaymiz
    s = NOISE_RE.sub(" ", s)

    # 4) Ortiqcha belgilar va bo'shliqlar
    s = re.sub(r"[,;:()\[\]]+", " ", s)
    s = re.sub(r"\s*-\s*", "-", s)                 # "Yangi - hayot" -> "Yangi-hayot"
    s = re.sub(r"\s+", " ", s).strip(" .-")
    s = s.replace(PH, APOSTROPHE)

    if not s:
        return ""
    if UPPERCASE:
        s = s.upper()
    return f'"{s}" {SUFFIX}'.strip()


# ---------------- Excel bilan ishlash ----------------
def pick_input(path):
    if path and os.path.exists(path):
        return path
    files = [f for f in os.listdir(".")
             if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
             and "_mahalla" not in f and "_xarita" not in f]
    if not files:
        sys.exit("Papkada .xlsx fayl topilmadi. Skriptni Excel fayl bilan bir papkaga qo'ying.")
    if len(files) == 1:
        return files[0]
    print("Qaysi faylni ishlaymiz?")
    for n, f in enumerate(files, 1):
        print(f"  {n}. {f}")
    while True:
        ch = input("Raqamni kiriting: ").strip()
        if ch.isdigit() and 1 <= int(ch) <= len(files):
            return files[int(ch) - 1]


def find_header(ws):
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip().lower() == SOURCE_HEADER:
                return r, c
    return None, None


def insert_column_after(ws, src_col, hdr_row):
    """src_col dan keyin yangi ustun qo'shadi; kenglik, stil, filtr, birlashgan kataklarni suradi."""
    new_col = src_col + 1

    # birlashtirilgan kataklarni vaqtincha ajratib, keyin surib qayta birlashtiramiz
    merged = [CellRange(str(m)) for m in ws.merged_cells.ranges]
    for m in merged:
        ws.unmerge_cells(str(m))

    # ustun kengliklarini eslab qolamiz
    widths = {}
    for letter, dim in list(ws.column_dimensions.items()):
        if dim.width:
            from openpyxl.utils import column_index_from_string
            widths[column_index_from_string(letter)] = dim.width

    ws.insert_cols(new_col)

    for idx in sorted(widths, reverse=True):
        if idx >= new_col:
            ws.column_dimensions[get_column_letter(idx + 1)].width = widths[idx]
    ws.column_dimensions[get_column_letter(new_col)].width = max(widths.get(src_col, 18), 22)

    for m in merged:
        if m.min_col >= new_col:
            m.shift(col_shift=1)
        elif m.max_col >= new_col:
            m.expand(right=1)
        ws.merge_cells(m.coord)

    # avtofiltr
    if ws.auto_filter and ws.auto_filter.ref:
        cr = CellRange(ws.auto_filter.ref)
        if cr.max_col >= src_col:
            cr.expand(right=1)
        ws.auto_filter.ref = cr.coord

    # "Mahalla" ustunining stilini nusxalaymiz
    for r in range(1, ws.max_row + 1):
        s, d = ws.cell(r, src_col), ws.cell(r, new_col)
        if s.has_style:
            d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
            d.alignment, d.number_format, d.protection = copy(s.alignment), s.number_format, copy(s.protection)
    ws.cell(hdr_row, new_col).value = NEW_HEADER
    return new_col


def similar_groups(names):
    """Yozilishi juda o'xshash nomlar juftligi (qo'lda tekshirish uchun)."""
    keys = sorted(names)
    pairs = []
    for i, a in enumerate(keys):
        for b in keys[i + 1:]:
            if abs(len(a) - len(b)) > 3:
                continue
            r = SequenceMatcher(None, a, b).ratio()
            if r >= SIMILARITY:
                pairs.append((a, b, round(r, 2)))
    return pairs


def process_sheet(ws, stats):
    hdr_row, src_col = find_header(ws)
    if not hdr_row:
        return False
    nxt = ws.cell(hdr_row, src_col + 1).value
    if nxt is not None and str(nxt).strip().lower() == NEW_HEADER.lower():
        new_col = src_col + 1
        print(f"  [{ws.title}] '{NEW_HEADER}' ustuni bor edi — yangilanadi")
    else:
        new_col = insert_column_after(ws, src_col, hdr_row)
        print(f"  [{ws.title}] '{NEW_HEADER}' ustuni {get_column_letter(new_col)} ga qo'shildi")

    for r in range(hdr_row + 1, ws.max_row + 1):
        raw = ws.cell(r, src_col).value
        res = normalize(raw)
        ws.cell(r, new_col).value = res or None
        if res:
            stats.setdefault(res, {"soni": 0, "variantlar": set()})
            stats[res]["soni"] += 1
            stats[res]["variantlar"].add(str(raw).strip())
        elif raw not in (None, ""):
            stats.setdefault("(bo'sh/aniqlanmadi)", {"soni": 0, "variantlar": set()})
            stats["(bo'sh/aniqlanmadi)"]["soni"] += 1
            stats["(bo'sh/aniqlanmadi)"]["variantlar"].add(str(raw).strip())
    return True


def write_report(wb, stats):
    name = "Tekshiruv"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.append(["Mahalla nomi", "Soni", "Asl yozilishlari"])
    for k in sorted(stats):
        v = stats[k]
        ws.append([k, v["soni"], " | ".join(sorted(v["variantlar"]))])
    ws.append([])
    ws.append(["O'xshash nomlar (bir mahalla bo'lishi mumkin — tekshiring)", "", ""])
    ws.append(["1-nom", "2-nom", "O'xshashlik"])
    pairs = similar_groups([k for k in stats if not k.startswith("(")])
    for a, b, r in pairs:
        ws.append([a, b, r])
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 80
    return pairs


def main():
    inp = pick_input(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT)
    out = os.path.splitext(inp)[0] + "_mahalla.xlsx"
    print(f"Ochilmoqda: {inp}")
    wb = load_workbook(inp)

    stats, done = {}, False
    for ws in wb.worksheets:
        if ws.title == "Tekshiruv":
            continue
        done |= process_sheet(ws, stats)
    if not done:
        sys.exit(f"'{SOURCE_HEADER}' sarlavhali ustun hech bir varaqda topilmadi.")

    pairs = write_report(wb, stats)
    try:
        wb.save(out)
    except PermissionError:
        sys.exit(f"'{out}' ochiq turibdi — Excel'da yopib, qayta ishga tushiring.")

    total = sum(v["soni"] for k, v in stats.items() if not k.startswith("("))
    print(f"\nSaqlandi: {out}")
    print(f"  {total} ta qator, {len([k for k in stats if not k.startswith('(')])} ta noyob mahalla")
    if pairs:
        print(f"  {len(pairs)} ta o'xshash nom juftligi topildi — 'Tekshiruv' varag'ini ko'ring")


if __name__ == "__main__":
    main()
