"""
TEZKOR: NAME_S ustunidan tashkiliy shaklni (MChJ, XK, OK, FXU, A/O, AOOT, MP,
MSG ...) olib tashlaydi, korxonaning yuridik nomining O'ZINI qoldiradi.
Katta fayl (100 mingdan ortiq qator) uchun: read_only + write_only.

Nom oldida ham, orqasida ham, lotin/kirillda ham ishlaydi; qo'shtirnoq ichidagi
nomga tegmaydi. Qiymat turlari (INN/SOATO raqamlari) saqlanadi.

Natija (yangi fayl, asl fayl saqlanadi): "NAME_S (tozalangan)" ustuni qo'shiladi
(IN_PLACE=True bo'lsa asl ustunning o'zi tozalanadi).

O'rnatish:  pip install openpyxl
"""
import os
import re
import sys
import time
import traceback
import unicodedata
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = "Книга1.xlsx"
OUTPUT_NAME = "NAME_tozalangan.xlsx"
NAME_HEADER_HINTS = ["name_s", "name", "nomi", "korxona", "tashkilot"]
IN_PLACE = False            # True: yangi ustun qo'shmay, asl ustunni tozalaydi
REMOVE_QUOTES = True        # True: natijada qo'shtirnoqlar ham olib tashlanadi

FORM_TOKENS = {
    "MCHJ", "XK", "OK", "FX", "FXU", "AJ", "AO", "AOOT", "OAJ", "QK", "DUK",
    "UK", "MP", "MSG", "OOO", "OAO", "CHP", "YATT", "DK", "XF",
}
FORM_PHRASES = [
    r"mas['`’ʼ‘]?uliyati\s+cheklangan\s+jamiyat\w*",
    r"xususiy\s+korxona\w*", r"xususiy\s+firma\w*", r"xorijiy\s+korxona\w*",
    r"oilaviy\s+korxona\w*", r"fermer\s+xo['`’ʼ‘]?jalig\w*",
    r"aksiyadorlik\s+jamiyat\w*", r"ochiq\s+aksiyadorlik\s+jamiyat\w*",
    r"qo['`’ʼ‘]?shma\s+korxona\w*", r"davlat\s+unitar\s+korxona\w*",
    r"масъулияти\s+чекланган\s+жамият\w*", r"хусусий\s+корхона\w*",
    r"оилавий\s+корхона\w*", r"фермер\s+хўжалиг\w*", r"хорижий\s+корхона\w*",
    r"акциядорлик\s+жамият\w*", r"хусусий\s+фирма\w*",
]
# ======================================================

if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]

CYR = {"М": "M", "Ч": "CH", "Ж": "J", "Х": "X", "К": "K", "О": "O", "Ф": "F",
       "У": "U", "А": "A", "П": "P", "С": "S", "Қ": "Q", "Д": "D", "Р": "R",
       "Н": "N", "Г": "G", "Т": "T", "В": "V", "Б": "B", "Э": "E", "И": "I",
       "Й": "Y", "Л": "L", "Е": "E", "З": "Z", "Ш": "SH", "Ў": "O", "Ғ": "G",
       "Ҳ": "H", "Ц": "TS", "Ю": "YU", "Я": "YA"}
PHRASE_RE = re.compile("|".join(FORM_PHRASES), re.I)
QUOTED = re.compile(r"[\"“”«»][^\"“”«»]*[\"“”«»]")
TRIM = re.compile(r"[.\-/'`’ʼ‘\s]")
FORMS_NORM = {t.upper() for t in FORM_TOKENS}
_cache = {}


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def tok_norm(t):
    s = "".join(CYR.get(ch, ch) for ch in t.upper())
    return TRIM.sub("", s)


def clean_name(value):
    s = nfc(value).strip()
    if not s:
        return s
    cached = _cache.get(s)
    if cached is not None:
        return cached
    saved = []

    def keep(m):
        saved.append(m.group(0))
        return f"\x00{len(saved) - 1}\x00"

    protected = QUOTED.sub(keep, s)
    protected = PHRASE_RE.sub(" ", protected)
    out = []
    for t in protected.split():
        if t.startswith("\x00"):
            out.append(t)
        elif tok_norm(t) not in FORMS_NORM:
            out.append(t)
    res = " ".join(out)
    res = re.sub(r"\x00(\d+)\x00", lambda m: saved[int(m.group(1))], res)
    if REMOVE_QUOTES:
        res = re.sub(r"[\"“”«»]", "", res)
    res = re.sub(r"\(\s*\)|«\s*»|\"\s*\"", " ", res)
    res = re.sub(r"\s+", " ", res).strip(" ,.-—()")
    if len(_cache) < 200000:
        _cache[s] = res
    return res


def styled_header(ws, values):
    out = []
    for v in values:
        c = WriteOnlyCell(ws, value=v)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center")
        out.append(c)
    return out


def main():
    path = INPUT_FILE if os.path.isabs(INPUT_FILE) else os.path.join(BASE_DIR, INPUT_FILE)
    if not os.path.exists(path):
        print("Fayl topilmadi:", path, "\nPapkadagi Excel fayllar:")
        for f in sorted(os.listdir(BASE_DIR)):
            if f.lower().endswith((".xlsx", ".xlsm")):
                print("  ", f)
        return
    print("Fayl:", os.path.basename(path))
    t0 = time.time()

    wb_in = load_workbook(path, read_only=True)
    ws_in = wb_in.active
    rows = ws_in.iter_rows(values_only=True)
    header = list(next(rows))
    ncol = len(header)

    name_idx, best = 0, None
    for j, v in enumerate(header):
        t = nfc(v).strip().lower()
        for i, hint in enumerate(NAME_HEADER_HINTS):
            if hint in t and (best is None or i < best[0]):
                best = (i, j)
    if best:
        name_idx = best[1]
    print(f"Nom ustuni: {get_column_letter(name_idx + 1)} ({header[name_idx]})")

    out_header = list(header)
    if not IN_PLACE:
        out_header.append("NAME_S (tozalangan)")

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Natija")
    ws.append(styled_header(ws, out_header))

    total = changed = 0
    for row in rows:
        row = list(row)
        if len(row) < ncol:
            row += [None] * (ncol - len(row))
        orig = row[name_idx]
        new = clean_name(orig)
        if IN_PLACE:
            row[name_idx] = new
        else:
            row.append(new)
        if new != nfc(orig).strip():
            changed += 1
        total += 1
        ws.append(row)

    out_ncol = ncol if IN_PLACE else ncol + 1
    name_out = name_idx if IN_PLACE else out_ncol - 1
    for j in range(out_ncol):
        L = get_column_letter(j + 1)
        ws.column_dimensions[L].width = 40 if j in (name_idx, name_out) else 14
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(out_ncol)}{total + 1}"

    out = os.path.join(BASE_DIR, OUTPUT_NAME)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"NAME_tozalangan_{datetime.now():%H%M%S}.xlsx")
        print("Natija fayli ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)
    wb_in.close()

    print(f"Jami: {total} | shakl olib tashlandi: {changed} | "
          f"o'zgarmadi: {total - changed}")
    print(f"Vaqt: {time.time() - t0:.1f} s")
    print("Saqlandi:", out)


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
