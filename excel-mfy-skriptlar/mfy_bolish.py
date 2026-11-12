# -*- coding: utf-8 -*-
"""
Papkadagi Excel fayllarni MAHALLA kesimida bo'lib chiqadi.
Har bir mahalla uchun alohida papka, ichida har bir fayldan o'sha mahallaga tegishli qatorlar.

Natija:
  MFY_kesimida/
    A.JOMIY MFY/
        A.JOMIY MFY - Yatt o'tmagan 11.11.2026 12-00.xlsx
        A.JOMIY MFY - ЮРИДИК_утмаган_....xlsx
    MFY aniqlanmagan/      <- mahallasi bo'sh yoki 1722210 kabi kod
    Svod.xlsx

Ishga tushirish:  py main.py
"""

import os
import re
import shutil
import sys
import time
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ================== SOZLAMALAR ==================
FILES = []                    # bo'sh — papkadagi barcha .xlsx
OUT_DIR = "MFY_kesimida"
MFY_HEADERS = ["mfy nomi (ro'yxat)", "mahalla nomi", "mfy nomi", "мфй номи", "mahalla", "махалла"]
MFY_WORDS = ["mfy", "мфй", "mahalla", "махалла", "маҳалла", "m.f.y", "м.ф.й"]
SKIP_WORDS = ["inn", "инн", "stir", "стир", "kod", "код", "soato", "соато", "id"]
NO_MFY = "MFY aniqlanmagan"
# ================================================

APOS = "'`‘’ʻʼ´"
QUOTES = "\"«»“”„"
PH = "\u0001"
NOISE_RE = re.compile(
    r"(?<![\w" + PH + r"])(?:mahalla\s+fuqarolar(?:i)?\s+yig\W?ini|fuqarolar\s+yig\W?ini|"
    r"m\s*\.?\s*f\s*\.?\s*y\s*\.?(?:\s*i)?|м\s*\.?\s*ф\s*\.?\s*й\s*\.?|mahallasi|mahalla|"
    r"махалласи|маҳалласи|махалла|маҳалла)(?![\w" + PH + r"])", re.I)


def mfy_name(value):
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ").strip()
    if not s or re.fullmatch(r"[\d\s.,-]*", s):
        return ""
    s = re.sub(r"(?<=\w)[" + re.escape(APOS) + r"](?=\w)", PH, s)
    s = re.sub("[" + re.escape(QUOTES + APOS) + "]", " ", s)
    s = NOISE_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip(" .,-").replace(PH, "`")
    return f'"{s.upper()}" MFY' if s else ""


def safe_name(s, maxlen=80):
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s[:maxlen] or "nomsiz"


def norm_header(v):
    s = re.sub(r"\s+", " ", str(v or "")).strip().lower()
    return re.sub("[`‘’ʻʼ´]", "'", s)


def text_share(ws, hdr, col, limit=300):
    n = good = 0
    for r in range(hdr + 1, min(ws.max_row, hdr + limit) + 1):
        v = ws.cell(r, col).value
        if v in (None, ""):
            continue
        n += 1
        if mfy_name(v):
            good += 1
    return good / n if n else 0


def find_header(ws):
    rows = range(1, min(ws.max_row, 30) + 1)
    for r in rows:
        heads = {}
        for c in range(1, ws.max_column + 1):
            h = norm_header(ws.cell(r, c).value)
            if h:
                heads.setdefault(h, c)
        for h in MFY_HEADERS:
            if h in heads:
                return r, heads[h], ws.cell(r, heads[h]).value
    best = None
    for r in rows:
        for c in range(1, ws.max_column + 1):
            h = norm_header(ws.cell(r, c).value)
            if not h or not any(w in h for w in MFY_WORDS):
                continue
            if any(re.search(r"\b" + re.escape(w) + r"\b", h) for w in SKIP_WORDS):
                continue
            share = text_share(ws, r, c)
            if share > 0 and (best is None or share > best[0]):
                best = (share, r, c)
    if best:
        return best[1], best[2], ws.cell(best[1], best[2]).value
    return None, None, None


def ask_column(fname, ws):
    print(f"  ! '{fname}' da mahalla ustuni avtomatik topilmadi.")
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [(c, ws.cell(r, c).value) for c in range(1, ws.max_column + 1) if ws.cell(r, c).value]
        if len(vals) >= 3:
            print(f"  {r}-qatordagi sarlavhalar:")
            for c, v in vals:
                print(f"    {c}. {get_column_letter(c)} — {str(v)[:50]}")
            ch = input("  Mahalla ustuni raqamini kiriting (o'tkazish uchun Enter): ").strip()
            if ch.isdigit() and 1 <= int(ch) <= ws.max_column:
                return r, int(ch), ws.cell(r, int(ch)).value
            return None, None, None
    return None, None, None


def copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
    if src.hyperlink:
        dst.hyperlink = copy(src.hyperlink)


def write_part(ws, hdr, rows, path):
    wb = Workbook()
    o = wb.active
    o.title = ws.title[:31]
    maxc = ws.max_column
    for c in range(1, maxc + 1):
        L = get_column_letter(c)
        d = ws.column_dimensions[L]
        if d.width:
            o.column_dimensions[L].width = d.width
        if d.hidden:
            o.column_dimensions[L].hidden = True
    out_r = 0
    for r in list(range(1, hdr + 1)) + rows:
        out_r += 1
        for c in range(1, maxc + 1):
            copy_cell(ws.cell(r, c), o.cell(out_r, c))
        h = ws.row_dimensions[r].height
        if h:
            o.row_dimensions[out_r].height = h
    for m in ws.merged_cells.ranges:
        if m.max_row <= hdr:
            o.merge_cells(str(m))
    o.auto_filter.ref = f"A{hdr}:{get_column_letter(maxc)}{out_r}"
    o.freeze_panes = o.cell(hdr + 1, 1)
    wb.save(path)


def prepare_out_dir():
    """Eski natijani tozalaydi; tarmoq diskida o'chmasa ham xato bermaydi."""
    if os.path.isdir(OUT_DIR):
        shutil.rmtree(OUT_DIR, ignore_errors=True)
        for _ in range(10):
            if not os.path.exists(OUT_DIR):
                break
            time.sleep(0.5)
    for _ in range(10):
        try:
            os.makedirs(OUT_DIR, exist_ok=True)
            return
        except OSError:
            time.sleep(0.5)
    os.makedirs(OUT_DIR, exist_ok=True)


def save_safe(func, *args):
    """Fayl ochiq bo'lsa — to'xtamaydi, xabar beradi."""
    try:
        func(*args)
        return True
    except PermissionError:
        print(f"  ! yozilmadi (Excel'da ochiq bo'lishi mumkin): {args[-1]}")
        return False


def main():
    files = FILES or sorted(
        f for f in os.listdir(".")
        if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
    )
    if not files:
        sys.exit("Papkada Excel fayl topilmadi. Skriptni fayllar turgan papkaga qo'ying.")
    print("Topilgan fayllar:")
    for f in files:
        print(f"  - {f}")

    prepare_out_dir()

    svod = {}
    done_files = []

    for fname in files:
        print(f"\n{fname}")
        try:
            wb = load_workbook(fname)
        except Exception as e:
            print(f"  ! ochilmadi: {e}")
            continue

        ws = hdr = col = title = None
        for w in wb.worksheets:
            hdr, col, title = find_header(w)
            if hdr:
                ws = w
                break
        if ws is None:
            ws = max(wb.worksheets, key=lambda w: w.max_row)
            hdr, col, title = ask_column(fname, ws)
            if not hdr:
                print("  o'tkazib yuborildi")
                continue
        print(f"  varaq: {ws.title}, mahalla ustuni: {get_column_letter(col)} ('{title}'), "
              f"sarlavha: {hdr}-qator")

        groups = {}
        for r in range(hdr + 1, ws.max_row + 1):
            if all(ws.cell(r, c).value in (None, "") for c in range(1, ws.max_column + 1)):
                continue
            key = mfy_name(ws.cell(r, col).value) or NO_MFY
            groups.setdefault(key, []).append(r)

        base = os.path.splitext(fname)[0]
        for key, rows in groups.items():
            folder = os.path.join(OUT_DIR, safe_name(key))
            os.makedirs(folder, exist_ok=True)
            out = os.path.join(folder, safe_name(f"{safe_name(key, 40)} - {base}", 150) + ".xlsx")
            save_safe(write_part, ws, hdr, rows, out)
            svod.setdefault(key, {})[fname] = len(rows)

        done_files.append(fname)
        print(f"  {sum(len(v) for v in groups.values())} qator -> {len(groups)} ta mahalla")

    if not done_files:
        sys.exit("Hech bir fayl bo'linmadi.")

    sw = Workbook()
    s = sw.active
    s.title = "Svod"
    s.append(["№", "Mahalla"] + [os.path.splitext(f)[0] for f in done_files] + ["Jami"])
    for c in s[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
    keys = sorted(k for k in svod if k != NO_MFY) + ([NO_MFY] if NO_MFY in svod else [])
    for i, k in enumerate(keys, 1):
        vals = [svod[k].get(f, 0) for f in done_files]
        s.append([i, k] + vals + [sum(vals)])
    tot = [sum(svod[k].get(f, 0) for k in keys) for f in done_files]
    s.append(["", "JAMI"] + tot + [sum(tot)])
    for c in s[s.max_row]:
        c.font = Font(bold=True)
    s.column_dimensions["A"].width = 6
    s.column_dimensions["B"].width = 32
    for j in range(len(done_files) + 1):
        s.column_dimensions[get_column_letter(3 + j)].width = 25
    s.freeze_panes = "C2"
    save_safe(sw.save, os.path.join(OUT_DIR, "Svod.xlsx"))

    print(f"\nTayyor: {OUT_DIR}\\ — {len(keys)} ta mahalla papkasi, {len(done_files)} ta fayl bo'lindi")
    if NO_MFY in svod:
        print(f"  '{NO_MFY}' papkasida {sum(svod[NO_MFY].values())} qator (mahallasi bo'sh yoki kod)")
    print(f"  Mahallalar bo'yicha soni: {OUT_DIR}\\Svod.xlsx")


if __name__ == "__main__":
    main()