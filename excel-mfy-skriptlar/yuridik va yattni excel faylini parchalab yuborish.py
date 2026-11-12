# -*- coding: utf-8 -*-
"""
YaTT o'tmagan va YURIDIK o'tmagan fayllarini TUMAN va MAHALLA kesimida bo'lib chiqadi
(alohida papkaga). Papkadagi boshqa fayllar (masalan "Кушимча оригинал") olinmaydi.

Natija:
  Otmaganlar_kesimida/
    Denov tumani/
        A.Jomiy_mfy_yatt_12.xlsx
        A.Jomiy_mfy_yuridik_3.xlsx
        MFY aniqlanmagan_yatt_4.xlsx
    Svod.xlsx

Ishga tushirish:  py otmaganlar_bolish.py
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
OUT_DIR = "Otmaganlar_kesimida"
# Fayl nomida shu so'z bo'lsa -> o'sha turdagi fayl. Boshqa fayllar olinmaydi.
FILE_TYPES = [
    ("yatt", "yatt"),
    ("ятт", "yatt"),
    ("юридик", "yuridik"),
    ("yuridik", "yuridik"),
]
TUMAN_HEADERS = ["tuman/shahar", "tuman", "туман", "туман номи", "tuman nomi", "манзилл"]
# Birinchi topilgani asosiy; u bo'sh bo'lgan qatorda keyingisi ishlatiladi
MFY_HEADERS = ["mfy nomi (ro'yxat)", "mahalla nomi", "мфй", "mfy", "mfy nomi", "мфй номи",
               "mahalla", "махалла"]
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


def mfy_core(value):
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ").strip()
    if not s or re.fullmatch(r"[\d\s.,-]*", s):
        return ""
    s = re.sub(r"(?<=\w)[" + re.escape(APOS) + r"](?=\w)", PH, s)
    s = re.sub("[" + re.escape(QUOTES + APOS) + "]", " ", s)
    s = NOISE_RE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip(" .,-").replace(PH, "`")
    return s.upper()


def tuman_core(value):
    if value is None:
        return ""
    s = re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip(" .,")
    if not s or re.fullmatch(r"[\d\s.,-]*", s):
        return ""
    s = re.sub("[" + re.escape(APOS) + "]", "`", s).upper()
    city = bool(re.search(r"\b(SHAHRI|SHAHAR|SH|ШАҲРИ|ШАХРИ|Ш)\.?$", s))
    s = re.sub(r"\s*\b(SHAHRI|SHAHAR|SH|ШАҲРИ|ШАХРИ|Ш|TUMANI|TUMAN|T|ТУМАНИ|ТУМАН|Т)\.?$", "", s).strip()
    if not s:
        return ""
    return f"{s} SHAHRI" if city else f"{s} TUMANI"


def nice(s):
    words = []
    for w in s.split(" "):
        words.append(".".join(p[:1].upper() + p[1:].lower() for p in w.split(".")))
    r = " ".join(words)
    return re.sub(r"\b(Shahri|Tumani)\b", lambda m: m.group(0).lower(), r)


def safe_name(s, maxlen=100):
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s[:maxlen] or "nomsiz"


def norm_header(v):
    s = re.sub(r"\s+", " ", str(v or "")).strip().lower()
    return re.sub("[`‘’ʻʼ´]", "'", s)


def find_cols(ws):
    for r in range(1, min(ws.max_row, 20) + 1):
        heads = {}
        for c in range(1, ws.max_column + 1):
            h = norm_header(ws.cell(r, c).value)
            if h:
                heads.setdefault(h, c)
        mcols = []
        for h in MFY_HEADERS:
            if h in heads and heads[h] not in mcols:
                mcols.append(heads[h])
        if mcols:
            tc = next((heads[h] for h in TUMAN_HEADERS if h in heads), None)
            return r, tc, mcols
    return None, None, []


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


def file_type(fname):
    low = fname.lower()
    for word, typ in FILE_TYPES:
        if word in low:
            return typ
    return None


def main():
    files = []
    for f in sorted(os.listdir(".")):
        if not f.lower().endswith((".xlsx", ".xlsm")) or f.startswith("~$"):
            continue
        t = file_type(f)
        if t:
            files.append((f, t))
    if not files:
        sys.exit("Papkada YaTT yoki YURIDIK o'tmagan fayli topilmadi.")
    print("Bo'linadigan fayllar:")
    for f, t in files:
        print(f"  [{t}] {f}")

    groups = {}
    types = []
    for fname, typ in files:
        print(f"\n{fname}")
        try:
            wb = load_workbook(fname)
        except Exception as e:
            print(f"  ! ochilmadi: {e}")
            continue
        found = []
        for w in wb.worksheets:
            hdr, tc, mcols = find_cols(w)
            if hdr:
                found.append((w.max_row, w, hdr, tc, mcols))
        if found:
            _, ws, hdr, tcol, mcols = max(found, key=lambda x: x[0])
        else:
            ws = max(wb.worksheets, key=lambda w: w.max_row)
            print("  ! MFY ustuni avtomatik topilmadi.")
            hdr, mc = ask_column(ws, "MFY")
            if not mc:
                print("  o'tkazib yuborildi")
                continue
            mcols, tcol = [mc], None
        if not tcol:
            print("  ! Tuman ustuni avtomatik topilmadi.")
            _, tcol = ask_column(ws, "Tuman")

        print(f"  varaq: {ws.title}, sarlavha: {hdr}-qator")
        print(f"  tuman ustuni: {get_column_letter(tcol) + ' (' + str(ws.cell(hdr, tcol).value) + ')' if tcol else 'yo`q'}")
        print(f"  MFY ustuni:   " + ", ".join(f"{get_column_letter(c)} ({ws.cell(hdr, c).value})" for c in mcols))

        n = 0
        for r in range(hdr + 1, ws.max_row + 1):
            if all(ws.cell(r, c).value in (None, "") for c in range(1, ws.max_column + 1)):
                continue
            t = (tuman_core(ws.cell(r, tcol).value) if tcol else "") or NO_TUMAN.upper()
            m = ""
            for c in mcols:
                m = mfy_core(ws.cell(r, c).value)
                if m:
                    break
            m = m or NO_MFY.upper()
            slot = groups.setdefault(t, {}).setdefault(m, {})
            if typ not in slot:
                slot[typ] = (ws, hdr, [])
            slot[typ][2].append(r)
            n += 1
        if typ not in types:
            types.append(typ)
        print(f"  {n} qator")

    if not groups:
        sys.exit("Hech bir fayl bo'linmadi.")

    prepare_out_dir()

    svod = []
    for t in sorted(groups):
        t_name = safe_name(nice(t))
        t_dir = os.path.join(OUT_DIR, t_name)
        os.makedirs(t_dir, exist_ok=True)
        for m in sorted(groups[t], key=lambda x: (x == NO_MFY.upper(), x)):
            counts = {}
            for typ, (ws, hdr, rows) in groups[t][m].items():
                k = len(rows)
                counts[typ] = k
                if m == NO_MFY.upper():
                    fn = f"{NO_MFY}_{typ}_{k}.xlsx"
                else:
                    fn = f"{safe_name(nice(m), 80)}_mfy_{typ}_{k}.xlsx"
                save_safe(write_part, ws, hdr, rows, os.path.join(t_dir, fn))
            svod.append((t_name, NO_MFY if m == NO_MFY.upper() else nice(m), counts))
        tot = {ty: sum(c.get(ty, 0) for tn, _, c in svod if tn == t_name) for ty in types}
        print(f"  {t_name}: {len(groups[t])} ta mahalla, " + ", ".join(f"{ty} {v}" for ty, v in tot.items()))

    sw = Workbook()
    s = sw.active
    s.title = "Svod"
    titles = {"yatt": "YaTT o'tmagan", "yuridik": "Yuridik o'tmagan"}
    s.append(["№", "Tuman", "Mahalla"] + [titles.get(ty, ty) for ty in types] + ["Jami"])
    for c in s[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    def total_row(label, rows, color):
        vals = [sum(c.get(ty, 0) for _, _, c in rows) for ty in types]
        s.append(["", label, ""] + vals + [sum(vals)])
        for c in s[s.max_row]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor=color)

    i = 0
    for t_name in sorted({x[0] for x in svod}):
        rows = [x for x in svod if x[0] == t_name]
        for _, m_name, counts in rows:
            i += 1
            vals = [counts.get(ty, 0) for ty in types]
            s.append([i, t_name, m_name] + vals + [sum(vals)])
        total_row(f"{t_name} jami", rows, "FFF2CC")
    total_row("JAMI", svod, "C6EFCE")

    s.column_dimensions["A"].width = 6
    s.column_dimensions["B"].width = 24
    s.column_dimensions["C"].width = 30
    for j in range(len(types) + 1):
        s.column_dimensions[get_column_letter(4 + j)].width = 16
    s.freeze_panes = "A2"
    save_safe(sw.save, os.path.join(OUT_DIR, "Svod.xlsx"))

    print(f"\nTayyor: {OUT_DIR}\\ — {len(groups)} ta tuman, {len(svod)} ta mahalla")
    print(f"  Umumiy jadval: {OUT_DIR}\\Svod.xlsx")


if __name__ == "__main__":
    main()