# -*- coding: utf-8 -*-
"""
Excel faylni TUMAN va MAHALLA kesimida bo'lib chiqadi.

Natija:
  Tuman_MFY_kesimida/
    Termiz shahri/
        Alisher Navoiy_mfy_23.xlsx      <- 23 = shu mahalladagi YaTT soni
        Farxod_mfy_15.xlsx
        MFY aniqlanmagan_4.xlsx         <- mahallasi bo'sh yoki kod
    Denov tumani/
        ...
    Svod.xlsx                           <- tuman va mahalla bo'yicha soni

Ishga tushirish (skriptni Excel fayl turgan papkaga qo'ying):
  py -m pip install openpyxl
  py tuman_mfy_bolish.py
"""

import os
import re
import shutil
import sys
import time
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================== SOZLAMALAR ==================
INPUT_FILE = "Кушимча оригинал 12.11.2026.xlsx"   # topilmasa — papkadan tanlatadi
OUT_DIR = "Tuman_MFY_kesimida"
# Ustun sarlavhalari (birinchi topilgani olinadi, katta-kichik harf farqi yo'q)
TUMAN_HEADERS = ["tuman/shahar", "tuman", "туман", "туман номи", "tuman nomi", "манзилл", "манзил", "manzil"]
MFY_HEADERS = ["мфй", "mfy", "мфй номи", "mfy nomi (ro'yxat)", "mfy nomi", "mahalla nomi", "mahalla", "махалла"]
NO_MFY = "MFY aniqlanmagan"
NO_TUMAN = "Tuman aniqlanmagan"
# ================================================

APOS = "'`‘’ʻʼ´"
QUOTES = "\"«»“”„"
PH = "\u0001"
NOISE_RE = re.compile(
    r"(?<![\w" + PH + r"])(?:mahalla\s+fuqarolar(?:i)?\s+yig\W?ini|fuqarolar\s+yig\W?ini|"
    r"m\s*\.?\s*f\s*\.?\s*y\s*\.?(?:\s*i)?|м\s*\.?\s*ф\s*\.?\s*й\s*\.?|mahallasi|mahalla|"
    r"махалласи|маҳалласи|махалла|маҳалла)(?![\w" + PH + r"])", re.I)


# ---------------- Nomlarni tozalash ----------------
def mfy_core(value):
    """'"ALISHER NAVOIY" MFY' / 'm.f.y. Alisher Navoiy' -> 'ALISHER NAVOIY'. Aniqlanmasa ''."""
    if value is None:
        return ""
    s = str(value).replace(" ", " ").strip()
    if not s or re.fullmatch(r"[\d\s.,-]*", s):        # bo'sh, 0, 1722210 kabi kod
        return ""
    s = re.sub(r"(?<=\w)[" + re.escape(APOS) + r"](?=\w)", PH, s)
    s = re.sub("[" + re.escape(QUOTES + APOS) + "]", " ", s)
    s = NOISE_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip(" .,-").replace(PH, "`")
    return s.upper()


def tuman_core(value):
    """'TERMIZ SHAHRI' / 'Termiz sh.' -> 'TERMIZ SHAHRI'; 'Denov t.' -> 'DENOV TUMANI'."""
    if value is None:
        return ""
    s = re.sub(r"\s+", " ", str(value).replace(" ", " ")).strip(" .,")
    if not s or re.fullmatch(r"[\d\s.,-]*", s):
        return ""
    s = re.sub(r"[" + re.escape(APOS) + r"]", "`", s)
    s = re.sub(r"\b(sh|ш)\.?$", "shahri", s, flags=re.I)
    s = re.sub(r"\b(t|т)\.?$", "tumani", s, flags=re.I)
    s = re.sub(r"\bshahar$", "shahri", s, flags=re.I)
    return s.upper()


def nice(s):
    """'A.AVLONIY' -> 'A.Avloniy', 'TERMIZ SHAHRI' -> 'Termiz shahri'."""
    words = []
    for w in s.split(" "):
        parts = [p[:1].upper() + p[1:].lower() for p in w.split(".")]
        words.append(".".join(parts))
    r = " ".join(words)
    return re.sub(r"\b(Shahri|Tumani|Tuman|Шаҳри|Тумани)\b", lambda m: m.group(0).lower(), r)


def safe_name(s, maxlen=100):
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s[:maxlen] or "nomsiz"


def norm_header(v):
    s = re.sub(r"\s+", " ", str(v or "")).strip().lower()
    return re.sub("[`‘’ʻʼ´]", "'", s)


# ---------------- Excel ----------------
def find_cols(ws):
    """-> (sarlavha qatori, tuman ustuni, mfy ustuni) yoki (None, None, None)"""
    for r in range(1, min(ws.max_row, 20) + 1):
        heads = {}
        for c in range(1, ws.max_column + 1):
            h = norm_header(ws.cell(r, c).value)
            if h:
                heads.setdefault(h, c)
        tc = next((heads[h] for h in TUMAN_HEADERS if h in heads), None)
        mc = next((heads[h] for h in MFY_HEADERS if h in heads), None)
        if mc:
            return r, tc, mc
    return None, None, None


def ask_column(ws, what):
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [(c, ws.cell(r, c).value) for c in range(1, ws.max_column + 1) if ws.cell(r, c).value]
        if len(vals) >= 3:
            print(f"  {r}-qatordagi sarlavhalar:")
            for c, v in vals:
                print(f"    {c}. {get_column_letter(c)} — {str(v)[:50]}")
            ch = input(f"  {what} ustuni raqamini kiriting (yo'q bo'lsa Enter): ").strip()
            return r, (int(ch) if ch.isdigit() and 1 <= int(ch) <= ws.max_column else None)
    return 1, None


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
    try:
        func(*args)
        return True
    except PermissionError:
        print(f"  ! yozilmadi (Excel'da ochiq bo'lishi mumkin): {args[-1]}")
        return False


def pick_file():
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        return sys.argv[1]
    if os.path.exists(INPUT_FILE):
        return INPUT_FILE
    files = sorted(f for f in os.listdir(".")
                   if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"))
    if not files:
        sys.exit("Papkada Excel fayl topilmadi. Skriptni fayl turgan papkaga qo'ying.")
    if len(files) == 1:
        return files[0]
    print("Qaysi faylni bo'lamiz?")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    while True:
        ch = input("Raqamni kiriting: ").strip()
        if ch.isdigit() and 1 <= int(ch) <= len(files):
            return files[int(ch) - 1]


def main():
    fname = pick_file()
    print(f"Ochilmoqda: {fname}")
    wb = load_workbook(fname)

    found = []
    for w in wb.worksheets:
        hdr, tc, mc = find_cols(w)
        if hdr:
            found.append((w.max_row, w, hdr, tc, mc))
    if found:
        _, ws, hdr, tcol, mcol = max(found, key=lambda x: x[0])
    else:
        ws = max(wb.worksheets, key=lambda w: w.max_row)
        print("  ! MFY ustuni avtomatik topilmadi.")
        hdr, mcol = ask_column(ws, "MFY")
        if not mcol:
            sys.exit("MFY ustuni tanlanmadi.")
        tcol = None
    if not tcol:
        print("  ! Tuman ustuni avtomatik topilmadi.")
        _, tcol = ask_column(ws, "Tuman")

    print(f"  varaq: {ws.title}, sarlavha: {hdr}-qator")
    print(f"  tuman ustuni: {get_column_letter(tcol) + ' (' + str(ws.cell(hdr, tcol).value) + ')' if tcol else 'yo`q'}")
    print(f"  MFY ustuni:   {get_column_letter(mcol)} ({ws.cell(hdr, mcol).value})")

    groups = {}
    for r in range(hdr + 1, ws.max_row + 1):
        if all(ws.cell(r, c).value in (None, "") for c in range(1, ws.max_column + 1)):
            continue
        t = (tuman_core(ws.cell(r, tcol).value) if tcol else "") or NO_TUMAN.upper()
        m = mfy_core(ws.cell(r, mcol).value) or NO_MFY.upper()
        groups.setdefault(t, {}).setdefault(m, []).append(r)

    prepare_out_dir()

    svod = []
    total = 0
    for t in sorted(groups):
        t_name = safe_name(nice(t))
        t_dir = os.path.join(OUT_DIR, t_name)
        os.makedirs(t_dir, exist_ok=True)
        t_total = 0
        for m in sorted(groups[t]):
            rows = groups[t][m]
            n = len(rows)
            if m == NO_MFY.upper():
                fn = f"{NO_MFY}_{n}.xlsx"
            else:
                fn = f"{safe_name(nice(m), 80)}_mfy_{n}.xlsx"
            save_safe(write_part, ws, hdr, rows, os.path.join(t_dir, fn))
            svod.append((t_name, nice(m) if m != NO_MFY.upper() else NO_MFY, n))
            t_total += n
        total += t_total
        print(f"  {t_name}: {len(groups[t])} ta mahalla, {t_total} ta YaTT")

    sw = Workbook()
    s = sw.active
    s.title = "Svod"
    s.append(["№", "Tuman", "Mahalla", "YaTT soni"])
    for c in s[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(horizontal="center")
    i = 0
    cur = None
    sub = 0
    for t_name, m_name, n in svod:
        if cur is not None and t_name != cur:
            s.append(["", f"{cur} jami", "", sub])
            for c in s[s.max_row]:
                c.font = Font(bold=True)
                c.fill = PatternFill("solid", fgColor="FFF2CC")
            sub = 0
        cur = t_name
        i += 1
        s.append([i, t_name, m_name, n])
        sub += n
    if cur is not None:
        s.append(["", f"{cur} jami", "", sub])
        for c in s[s.max_row]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="FFF2CC")
    s.append(["", "VILOYAT JAMI", "", total])
    for c in s[s.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="C6EFCE")
    s.column_dimensions["A"].width = 6
    s.column_dimensions["B"].width = 26
    s.column_dimensions["C"].width = 32
    s.column_dimensions["D"].width = 12
    s.freeze_panes = "A2"
    save_safe(sw.save, os.path.join(OUT_DIR, "Svod.xlsx"))

    print(f"\nTayyor: {OUT_DIR}\\ — {len(groups)} ta tuman, {len(svod)} ta mahalla fayli, jami {total} ta YaTT")
    print(f"  Umumiy jadval: {OUT_DIR}\\Svod.xlsx")


if __name__ == "__main__":
    main()
