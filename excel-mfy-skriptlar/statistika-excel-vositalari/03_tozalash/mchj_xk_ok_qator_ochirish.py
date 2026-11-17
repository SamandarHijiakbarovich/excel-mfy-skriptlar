"""
TEZKOR: NAME_S da MCHJ, XK yoki OK bo'lgan qatorlarni jadvaldan o'chiradi.
Katta fayl (100 mingdan ortiq qator) uchun. Faylni bir marta oqib, kerakli
qatorlarni yangi faylga yozadi (read_only + write_only) - bu bittalab
o'chirishdan o'nlab marta tez.

Shakl TOKEN sifatida solishtiriladi (lotin/kirill, katta-kichik harf farqi yo'q),
qo'shtirnoq ichidagi nomga qaralmaydi. Qiymat turlari (raqam/matn) saqlanadi.

Natija (yangi fayl, asl fayl saqlanadi):
  * asosiy varaq: MCHJ/XK/OK qatorlari olib tashlangan;
  * "O'chirilganlar" varag'i: o'chirilgan qatorlar (tekshirish uchun).

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
INPUT_FILE = "Книга1.xlsx"                 # yoki aniq fayl nomi
OUTPUT_NAME = "MCHJ_XK_OK_olingan.xlsx"
NAME_HEADER_HINTS = ["name_s", "name", "nomi", "korxona", "tashkilot"]
REMOVE_FORMS = {"MCHJ", "XK", "OK"}        # shu shakllar bo'lgan qatorlar o'chadi
SAVE_REMOVED = True
# ======================================================

if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]

CYR = {"М": "M", "Ч": "CH", "Ж": "J", "Х": "X", "К": "K", "О": "O", "Ф": "F",
       "У": "U", "А": "A", "П": "P", "С": "S", "Қ": "Q", "Д": "D", "Р": "R",
       "Н": "N", "Г": "G", "Т": "T", "В": "V", "Б": "B", "Э": "E", "И": "I",
       "Й": "Y", "Л": "L", "Е": "E", "З": "Z", "Ш": "SH", "Ў": "O", "Ғ": "G",
       "Ҳ": "H", "Ц": "TS", "Ю": "YU", "Я": "YA"}
QUOTED = re.compile(r"[\"“”«»][^\"“”«»]*[\"“”«»]")
TRIM = re.compile(r"[.\-/'`’ʼ‘]")
_cache = {}


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def tok_norm(t):
    s = "".join(CYR.get(ch, ch) for ch in t.upper())
    return TRIM.sub("", s)


def has_form(value):
    """NAME_S da REMOVE_FORMS dan biri TOKEN sifatida bormi? -> topilgan shakl yoki None"""
    if value is None:
        return None
    s = nfc(value)
    if s in _cache:
        return _cache[s]
    res = None
    if s.strip():
        for t in QUOTED.sub(" ", s).split():
            if tok_norm(t) in REMOVE_FORMS:
                res = tok_norm(t)
                break
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

    # nom ustunini topamiz
    name_idx = 0
    best = None
    for j, v in enumerate(header):
        t = nfc(v).strip().lower()
        for i, hint in enumerate(NAME_HEADER_HINTS):
            if hint in t and (best is None or i < best[0]):
                best = (i, j)
    if best:
        name_idx = best[1]
    print(f"Nom ustuni: {get_column_letter(name_idx + 1)} ({header[name_idx]})")

    # natija fayli - write_only (tez)
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Natija")
    ws.append(styled_header(ws, header))
    sr = None
    if SAVE_REMOVED:
        sr = wb.create_sheet("O'chirilganlar")
        sr.append(styled_header(sr, header))

    kept = removed = 0
    counts = {f: 0 for f in REMOVE_FORMS}
    for row in rows:
        row = list(row)
        if len(row) < ncol:
            row += [None] * (ncol - len(row))
        val = row[name_idx] if name_idx < len(row) else None
        f = has_form(val)
        if f:
            counts[f] += 1
            removed += 1
            if sr is not None:
                sr.append(row)
        else:
            kept += 1
            ws.append(row)

    # ko'rinish: kenglik, muzlatish, filtr
    for s in (ws, sr) if sr else (ws,):
        for j in range(ncol):
            L = get_column_letter(j + 1)
            s.column_dimensions[L].width = 40 if j == name_idx else 14
        s.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncol)}{kept + 1}"

    out = os.path.join(BASE_DIR, OUTPUT_NAME)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"MCHJ_XK_OK_olingan_{datetime.now():%H%M%S}.xlsx")
        print("Natija fayli ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)
    wb_in.close()

    print(f"O'chirildi: {removed} qator  (" +
          ", ".join(f"{k}: {v}" for k, v in counts.items()) + ")")
    print(f"Qoldi: {kept} qator")
    print(f"Vaqt: {time.time() - t0:.1f} s")
    print("Saqlandi:", out)
    if SAVE_REMOVED and removed:
        print("O'chirilganlar \"O'chirilganlar\" varag'ida.")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
