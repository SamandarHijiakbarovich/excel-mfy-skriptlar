"""
Qo'mita bergan "Hashlangan" fayldagi SHA-256 ustunini KALIT yordamida ochadi.

Bu hashni "sindirmaydi". Ikki vakolatli manbani ulaydi (join):
  REJIM 1 - KALIT = tayyor jadval: kalit faylida 64-belgili hash ustuni va
            uning yonida asl qiymat(lar) bo'lsa -> to'g'ridan-to'g'ri ulaydi.
  REJIM 2 - KALIT = o'z bazangiz: faylda hash yo'q, lekin asl identifikator
            (JSHSHIR/INN) bor -> o'sha ustunni SHA-256 bilan qayta hashlab,
            qaysi ustun+ko'rinish Hashlangan fayldagi hashlarni hosil
            qilishini O'ZI topadi, so'ng ulaydi.

Natija: Hashlangan faylning nusxasi + "Ochilgan ..." ustunlari (asl qiymat va
kalitdagi boshqa maydonlar, masalan YaTT/yuridik turi). Asl ustunlarga tegilmaydi.
Oxirida mos kelmagan qatorlar va moslik foizi hisobot qilinadi.

O'rnatish:  pip install pandas openpyxl
Ishlatish:  python hash_ochish.py [hashlangan.xlsx] [kalit.xlsx]
"""
import hashlib
import os
import re
import sys
import traceback
import unicodedata
from copy import copy
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # yoki r"Z:\МОЛИЯ БУЛИМИ\1"
HASHED_FILE = "Hashlangan.xlsx"      # Qo'mita bergan fayl (A ustuni - hash)
KEY_FILE = "kalit.xlsx"              # kalit jadval YOKI o'z bazangiz
OUTPUT_NAME = "Ochilgan.xlsx"

HASH_ALGO = "sha256"                 # Qo'mita shu bilan hashlagan (odatda sha256)
# Qanday ko'rinishlarni sinab ko'rsin (REJIM 2 uchun). Qo'mita hashlashdan oldin
# qiymatni qanday tayyorlagan bo'lsa, shulardan biri mos keladi:
NORMALIZERS = ["raw", "strip", "digits", "upper", "lower", "nospace",
               "upper_nospace", "strip_digits"]
GET_EXTRA_FROM_KEY = True            # kalitdagi boshqa ustunlarni ham olib kelsin
HIGHLIGHT = True
# ======================================================

if len(sys.argv) > 1:
    HASHED_FILE = sys.argv[1]
if len(sys.argv) > 2:
    KEY_FILE = sys.argv[2]

HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
GREEN = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
RED = PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD")
HEAD_FILL = PatternFill("solid", start_color="375623", end_color="375623")
HEAD_FONT = Font(bold=True, color="FFFFFF")
NEW_FILL = PatternFill("solid", start_color="EEF6E8", end_color="EEF6E8")


def nfc(v):
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else \
        unicodedata.normalize("NFC", str(v))


def normalize(v, mode):
    s = nfc(v)
    if mode == "raw":
        return s
    if mode == "strip":
        return s.strip()
    if mode == "digits":
        return re.sub(r"\D", "", s)
    if mode == "strip_digits":
        return re.sub(r"\D", "", s.strip())
    if mode == "upper":
        return s.strip().upper()
    if mode == "lower":
        return s.strip().lower()
    if mode == "nospace":
        return re.sub(r"\s+", "", s)
    if mode == "upper_nospace":
        return re.sub(r"\s+", "", s).upper()
    return s


def h(s):
    return hashlib.new(HASH_ALGO, s.encode("utf-8")).hexdigest()


def resolve(name):
    p = name if os.path.isabs(name) else os.path.join(BASE_DIR, name)
    if os.path.exists(p):
        return p
    key = re.sub(r"[^a-z0-9]", "", name.lower())
    for f in sorted(os.listdir(BASE_DIR)):
        if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$") \
                and re.sub(r"[^a-z0-9]", "", f.lower()) == key:
            return os.path.join(BASE_DIR, f)
    return None


def pick_excel(title):
    files = [f for f in sorted(os.listdir(BASE_DIR))
             if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    print(title)
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    try:
        return os.path.join(BASE_DIR, files[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def find_hash_col(df):
    """64-belgili hex eng ko'p uchraydigan ustun indeksi."""
    best, best_share = None, 0.0
    for j in range(df.shape[1]):
        col = df.iloc[:, j].map(lambda v: bool(HEX64.match(nfc(v).strip())))
        share = col.mean() if len(col) else 0
        if share > best_share:
            best, best_share = j, share
    return (best, best_share) if best_share >= 0.5 else (None, best_share)


def read_sheet_autohdr(path, sheet):
    """Sarlavha qatorini topib DataFrame qaytaradi."""
    head = pd.read_excel(path, sheet_name=sheet, header=None, nrows=12, dtype=str)
    hdr = 0
    for i in range(len(head)):
        vals = [nfc(x).strip() for x in head.iloc[i].tolist()]
        nonempty = [v for v in vals if v]
        hexish = sum(bool(HEX64.match(v)) for v in nonempty)
        if nonempty and hexish == 0 and len(nonempty) >= 2:
            hdr = i
            break
    return pd.read_excel(path, sheet_name=sheet, header=hdr, dtype=object)


# ---------------------------------------------------------- kalitni qurish
def build_from_mapping(df, hcol):
    """REJIM 1: kalit faylida hash ustuni bor."""
    other = [c for c in range(df.shape[1]) if c != hcol]
    mapping = {}
    for _, row in df.iterrows():
        hv = nfc(row.iloc[hcol]).strip().lower()
        if not HEX64.match(hv):
            continue
        mapping.setdefault(hv, {df.columns[c]: row.iloc[c] for c in other})
    cols = [df.columns[c] for c in other]
    return mapping, cols, "mapping"


def build_from_base(df, target_hashes):
    """REJIM 2: o'z bazangizdan SHA-256 ni qayta hosil qilib mos ustunni topadi."""
    best = None   # (share, col_idx, mode, mapping)
    for j in range(df.shape[1]):
        series = df.iloc[:, j]
        if series.map(lambda v: nfc(v).strip() == "").mean() > 0.6:
            continue
        for mode in NORMALIZERS:
            hashed = {}
            matched = 0
            for idx, v in series.items():
                s = normalize(v, mode)
                if not s:
                    continue
                hv = h(s)
                if hv in target_hashes:
                    matched += 1
                    hashed.setdefault(hv, idx)
            share = matched / max(1, len(target_hashes))
            if best is None or share > best[0]:
                best = (share, j, mode, hashed)
            if share >= 0.98:
                break
        if best and best[0] >= 0.98:
            break
    if not best or best[0] < 0.30:
        return None, None, None, best[0] if best else 0.0
    share, j, mode, hashed = best
    other = [c for c in range(df.shape[1])]   # bazadagi hamma ustun foydali
    mapping = {hv: {df.columns[c]: df.iloc[idx, c] for c in other}
               for hv, idx in hashed.items()}
    cols = [df.columns[c] for c in other]
    print(f"  Mos ustun: '{df.columns[j]}' | ko'rinish: {mode} | "
          f"moslik: {share:.1%}")
    return mapping, cols, f"base[{df.columns[j]}/{mode}]", share


def build_key(path, target_hashes):
    xl = pd.ExcelFile(path)
    for sheet in xl.sheet_names:
        df = read_sheet_autohdr(path, sheet)
        if df.empty:
            continue
        hcol, share = find_hash_col(df)
        if hcol is not None:
            print(f"  Kalit turi: tayyor jadval (hash ustuni = "
                  f"'{df.columns[hcol]}', {share:.0%})")
            return build_from_mapping(df, hcol)
        res = build_from_base(df, target_hashes)
        if res[0] is not None:
            print(f"  Kalit turi: o'z bazangiz (SHA-256 qayta hosil qilindi)")
            return res[0], res[1], res[2]
    return None, None, None


# ------------------------------------------------------------------ asosiy
def main():
    hpath = resolve(HASHED_FILE) or pick_excel("Hashlangan faylni tanlang:")
    kpath = resolve(KEY_FILE) or pick_excel("Kalit (yoki o'z bazangiz) faylni tanlang:")
    if not hpath or not kpath:
        print("Fayllar topilmadi.")
        return
    print("Hashlangan:", os.path.basename(hpath))
    print("Kalit     :", os.path.basename(kpath))

    wb = load_workbook(hpath)
    ws = wb.active
    # hash ustuni va sarlavha qatori
    probe = pd.read_excel(hpath, sheet_name=ws.title, header=None, dtype=str)
    hcol0, share = find_hash_col(probe)
    if hcol0 is None:
        print("Hashlangan faylda 64-belgili SHA-256 ustuni topilmadi.")
        return
    c_hash = hcol0 + 1
    # sarlavha bormi? birinchi qator hash bo'lsa - sarlavhasiz
    first = nfc(ws.cell(1, c_hash).value).strip()
    hdr = 0 if HEX64.match(first) else 1
    data_start = hdr + 1
    target = {nfc(ws.cell(r, c_hash).value).strip().lower()
              for r in range(data_start, ws.max_row + 1)
              if HEX64.match(nfc(ws.cell(r, c_hash).value).strip())}
    print(f"Hash ustuni: {get_column_letter(c_hash)} | {len(target)} ta noyob hash")

    print("Kalit o'qilmoqda va moslik aniqlanmoqda...")
    mapping, extra_cols, mode = build_key(kpath, target)
    if not mapping:
        print("MOS KELMADI. Ehtimol Qo'mita salt (tuz) qo'shgan yoki boshqa "
              "algoritm ishlatgan. Bu holda Qo'mitadan 'hash -> asl qiymat' "
              "jadvalini (tayyor kalit) so'rash kerak.")
        return

    # natija ustunlari
    last = max(c.column for c in ws[max(hdr, 1)] if c.value not in (None, "")) \
        if hdr else ws.max_column
    out_cols = []
    use_cols = extra_cols if GET_EXTRA_FROM_KEY else extra_cols[:1]
    for name in use_cols:
        last += 1
        out_cols.append((last, name))
        cell = ws.cell(max(hdr, 1), last)
        cell.value = f"Ochilgan: {name}"

    matched = 0
    for r in range(data_start, ws.max_row + 1):
        hv = nfc(ws.cell(r, c_hash).value).strip().lower()
        rec = mapping.get(hv)
        if rec:
            matched += 1
            for col, name in out_cols:
                ws.cell(r, col).value = rec.get(name)
        if HIGHLIGHT:
            fill = GREEN if rec else RED
            ws.cell(r, c_hash).fill = fill

    # bezak: yangi ustunlar sarlavhasi
    if hdr:
        for col, _ in out_cols:
            c = ws.cell(1, col)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
        for r in range(2, ws.max_row + 1):
            for col, _ in out_cols:
                ws.cell(r, col).fill = NEW_FILL
        for col, _ in out_cols:
            ws.column_dimensions[get_column_letter(col)].width = 22

    out = os.path.join(BASE_DIR, OUTPUT_NAME)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"Ochilgan_{datetime.now():%H%M%S}.xlsx")
        print("Natija fayli ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)

    total = ws.max_row - data_start + 1
    print(f"Jami: {total} | ochildi: {matched} | mos kelmadi: {total - matched} "
          f"| {matched / max(1, total):.1%}")
    print("Usul:", mode)
    print("Saqlandi:", out)
    print("Yashil = ochildi, qizil = kalitda topilmadi (A ustunida belgilangan).")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
