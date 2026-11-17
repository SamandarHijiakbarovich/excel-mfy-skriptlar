"""
Faylni avval TUMAN papkalariga, keyin har papka ichida MAHALLA bo'yicha fayllarga
bo'ladi. Fayl nomi = mahalla nomi. Har bir mahalla faylida ikkala manba varaq
("Сўровномалар" va "Интервьюерлар") o'z nomi bilan alohida varaq, oxirida soni.

Papka ko'rinishi:
  <fayl>_tuman_mahalla/
      Termiz tumani/
          Mustaqillik.xlsx
          Oq machit.xlsx
          00_Jami_royxat.xlsx   (shu tuman mahallari bo'yicha)
      Denov tumani/
          ...

Har varaqdagi TUMAN va MAHALLA ustunini o'zi topadi. Viloyat ustuni bo'lsa,
faqat Surxondaryo olinadi.

O'rnatish:  pip install openpyxl
Ishlatish:  python mahalla_kesimida_bolish.py [fayl.xlsx]
"""
import os
import re
import shutil
import sys
import time
import traceback
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Bir nechta faylni birdan bo'lish mumkin. Bo'sh qoldirsangiz, papkadagi
# BARCHA Excel fayllar (natija papkalaridan tashqari) bo'linadi.
INPUT_FILES = []   # masalan: ["yatt_royxati.xlsx", "Yuridik.xlsx"]
OUTPUT_DIR = None                  # None -> "<fayl nomi>_tuman_mahalla"
REGION_NAME = r"sur\w*dar"
ONLY_REGION = True                 # False -> hamma viloyat
STRIP_MFY = True                   # fayl nomidan "m.f.y." prefiksini olib tashlash
# Varaqdan olib tashlanadigan ustunlar (Excel harflari). Kalit = varaq nomining
# bir qismi (kichik harf). Masalan Сўровномалар dan B, P..AA ni tashlash:
DROP_COLS = {
    "сўровнома": ["B", "E", "I", "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"],
    "интервью": [],
}
COUNT_LABEL = "Jami:"
# ======================================================

if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]

CYR = {"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
       "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
       "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
       "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "",
       "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "ғ": "g",
       "қ": "q", "ҳ": "h"}
REGION_RE = re.compile(REGION_NAME, re.I)
MFY_TOKEN = re.compile(
    r"\b(m\.?\s*f\.?\s*y\.?|mfy|мфй|м\.?\s*ф\.?\s*й\.?|mahallasi|mahalla|"
    r"maxalla|маҳалла\w*|махалла\w*|msg|мсг|qfy|кфй)\b", re.I)
QUOTES_RE = re.compile(r"[\"«»“”`'’ʼ]")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def lat(v):
    return "".join(CYR.get(ch, ch) for ch in nfc(v).casefold())


def clean_mah(v):
    """Mahalla nomidan m.f.y./mahalla/qo'shtirnoqlarni tozalaydi."""
    s = QUOTES_RE.sub("", nfc(v).strip())
    s = MFY_TOKEN.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip(" -.,")


def norm_key(v):
    """Guruhlash kaliti - imlo/joylashuv farqini birlashtiradi."""
    return re.sub(r"[^a-z0-9]", "", lat(clean_mah(v)))


def disp_name(v):
    s = clean_mah(v) if STRIP_MFY else nfc(v).strip()
    return s or nfc(v).strip()


def safe_filename(s):
    s = nfc(s).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    return re.sub(r"\s+", " ", s).strip(" .") or "nomsiz"


def drop_indices(title):
    """Shu varaq uchun tashlanadigan 0-indeksli ustunlar to'plami."""
    out = set()
    lt = lat(title)
    for keyword, letters in DROP_COLS.items():
        if lat(keyword) in lt:
            for L in letters:
                try:
                    out.add(column_index_from_string(str(L).strip().upper()) - 1)
                except Exception:
                    pass
    return out


def read_header(ws):
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return list(row)
    return []


def find_cols(header):
    """-> (tuman, mahalla, viloyat) ustun indekslari."""
    tcol = mcol = vcol = None
    for j, v in enumerate(header):
        h = lat(v).strip()
        if not h:
            continue
        if vcol is None and ("hudud" in h or "viloyat" in h):
            vcol = j
        if tcol is None and "tuman" in h and "viloyat" not in h and "hudud" not in h:
            tcol = j
        if mcol is None and ("mahalla" in h or "maxalla" in h):
            mcol = j
    return tcol, mcol, vcol


def styled_header(ws, values, fill):
    out = []
    for v in values:
        c = WriteOnlyCell(ws, value=v)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=fill, end_color=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        out.append(c)
    return out


def count_cells(ws, ncol, label):
    cells = []
    for j in range(ncol):
        c = WriteOnlyCell(ws, value=label if j == 0 else None)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
        c.border = BOX
        c.alignment = Alignment(horizontal="left" if j == 0 else "center")
        cells.append(c)
    return cells


def process_one(path):
    print("\n=== Fayl:", os.path.basename(path), "===")
    t0 = time.time()

    wb = load_workbook(path, read_only=True)
    sheets = []
    for ws in wb.worksheets:
        header = read_header(ws)
        if not header:
            continue
        tcol, mcol, vcol = find_cols(header)
        if (tcol is not None and mcol is not None):
            pass
        if tcol is None or mcol is None:
            print(f"  [{ws.title}] Туман yoki Маҳалла ustuni yo'q - o'tkazildi")
            continue
        drop = drop_indices(ws.title)
        sheets.append((ws.title, header, tcol, mcol, vcol, drop))
        print(f"  [{ws.title}] Туман={get_column_letter(tcol + 1)}, "
              f"Маҳалла={get_column_letter(mcol + 1)}"
              + (f", viloyat={get_column_letter(vcol + 1)}" if vcol is not None else ""))
    if not sheets:
        print("Hech bir varaqda Туман+Маҳалла ustuni topilmadi.")
        return

    # data[sheet][tuman_key][mahalla_key] = [rows]
    data = {t: defaultdict(lambda: defaultdict(list)) for t, *_ in sheets}
    tuman_disp, mah_disp = {}, {}
    ncols = {}
    skipped = Counter()
    for title, header, tcol, mcol, vcol, drop in sheets:
        ws = wb[title]
        ncols[title] = len(header)
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            row = list(row)
            if ONLY_REGION and vcol is not None:
                if not REGION_RE.search(lat(row[vcol] if vcol < len(row) else "")):
                    skipped[title] += 1
                    continue
            tv = row[tcol] if tcol < len(row) else None
            mv = row[mcol] if mcol < len(row) else None
            tk = " ".join(lat(tv).split()) or "__tuman_yoq__"
            mk = norm_key(mv) or "__mahalla_yoq__"
            tuman_disp.setdefault(tk, nfc(tv).strip() or "Tuman ko'rsatilmagan")
            mah_disp.setdefault((tk, mk), disp_name(mv) or "Mahalla ko'rsatilmagan")
            data[title][tk][mk].append(row)

    wb.close()

    out_dir = OUTPUT_DIR or os.path.join(
        BASE_DIR, os.path.splitext(os.path.basename(path))[0] + "_tuman_mahalla")
    # eski natijani tozalaymiz (takror (2),(3) chiqmasligi uchun)
    if os.path.isdir(out_dir):
        try:
            shutil.rmtree(out_dir)
            print("Eski papka tozalandi.")
        except Exception as e:
            print("Eski papkani tozalab bo'lmadi:", e)
    os.makedirs(out_dir, exist_ok=True)
    palette = ["1F4E78", "375623", "833C00", "7030A0"]

    drops = {t: d for (t, _h, _tc, _mc, _vc, d) in sheets}
    all_tumans = sorted(set().union(*[set(data[t]) for t, *_ in sheets]),
                        key=lambda k: tuman_disp.get(k, k).lower())
    grand = Counter()
    n_files = 0
    for tk in all_tumans:
        tdir = os.path.join(out_dir, safe_filename(tuman_disp[tk]))
        os.makedirs(tdir, exist_ok=True)
        mah_keys = sorted(
            set().union(*[set(data[t][tk]) for t, *_ in sheets]),
            key=lambda k: mah_disp.get((tk, k), k).lower())
        rows_for_index = []
        for mk in mah_keys:
            name = mah_disp[(tk, mk)]
            nb = Workbook(write_only=True)
            counts = []
            for si, (title, header, tcol, mcol, vcol, drop) in enumerate(sheets):
                rows = data[title][tk].get(mk, [])
                counts.append(len(rows))
                if not rows:
                    continue
                ws = nb.create_sheet(title[:31])
                nc = ncols[title]
                keep = [j for j in range(nc) if j not in drop]
                out_header = [header[j] if j < len(header) else None for j in keep]
                ws.append(styled_header(ws, out_header, palette[si % len(palette)]))
                for row in rows:
                    full = list(row) + [None] * (nc - len(row))
                    ws.append([full[j] for j in keep])
                ws.append(count_cells(ws, len(keep), f"{COUNT_LABEL} {len(rows)} ta"))
                for j in range(len(keep)):
                    ws.column_dimensions[get_column_letter(j + 1)].width = 18
                ws.freeze_panes = "A2"
            if not nb.sheetnames:
                continue
            fname = safe_filename(name) + ".xlsx"
            # bir tumanda bir xil fayl nomi takrorlansa
            fp = os.path.join(tdir, fname)
            k = 2
            while os.path.exists(fp):
                fp = os.path.join(tdir, f"{safe_filename(name)} ({k}).xlsx")
                k += 1
            nb.save(fp)
            n_files += 1
            rows_for_index.append((name, counts))
            for i, c in enumerate(counts):
                grand[sheets[i][0]] += c

        # shu tuman uchun mahalla svodkasi
        ib = Workbook()
        iws = ib.active
        iws.title = "Mahallalar"
        head = ["Mahalla"] + [t for t, *_ in sheets] + ["Jami"]
        iws.append(head)
        for c in iws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
        tot = [0] * len(sheets)
        for name, counts in sorted(rows_for_index, key=lambda x: -sum(x[1])):
            iws.append([name] + counts + [sum(counts)])
            for i, c in enumerate(counts):
                tot[i] += c
            for c in iws[iws.max_row]:
                c.border = BOX
        iws.append(["JAMI"] + tot + [sum(tot)])
        for c in iws[iws.max_row]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
            c.border = BOX
        iws.column_dimensions["A"].width = 30
        for j in range(2, len(head) + 1):
            iws.column_dimensions[get_column_letter(j)].width = 16
        iws.freeze_panes = "A2"
        iws.auto_filter.ref = f"A1:{get_column_letter(len(head))}{iws.max_row}"
        ib.save(os.path.join(tdir, "00_Jami_royxat.xlsx"))
        print(f"  {tuman_disp[tk]:26} | {len(rows_for_index)} ta mahalla")

    print(f"Jami: {len(all_tumans)} ta tuman, {n_files} ta mahalla fayli")
    for t, *_ in sheets:
        print(f"  {t}: {grand[t]} ta" +
              (f" (boshqa viloyat o'tkazildi: {skipped[t]})" if skipped[t] else ""))
    print(f"Vaqt: {time.time() - t0:.1f} s")
    print("Papka:", out_dir)


def collect_files():
    import sys
    if len(sys.argv) > 1:
        return [a if os.path.isabs(a) else os.path.join(BASE_DIR, a) for a in sys.argv[1:]]
    if INPUT_FILES:
        return [f if os.path.isabs(f) else os.path.join(BASE_DIR, f) for f in INPUT_FILES]
    # papkadagi barcha Excel fayllar (natija/jami fayllaridan tashqari)
    out = []
    for f in sorted(os.listdir(BASE_DIR)):
        if (f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")
                and not f.startswith("00_") and "natija" not in f.lower()):
            out.append(os.path.join(BASE_DIR, f))
    return out


def main():
    files = collect_files()
    files = [f for f in files if os.path.exists(f)]
    if not files:
        print("Bo'linadigan Excel fayl topilmadi.")
        print("Papkadagi fayllar:")
        for f in sorted(os.listdir(BASE_DIR)):
            if f.lower().endswith((".xlsx", ".xlsm")):
                print("  ", f)
        return
    print(f"Jami {len(files)} ta fayl bo'linadi:")
    for f in files:
        print("  -", os.path.basename(f))
    for f in files:
        try:
            process_one(f)
        except PermissionError as e:
            print("  Faylga kirib bo'lmadi (Excelda ochiq bo'lsa yoping):", e)
        except Exception:
            print("  Shu faylda xato:")
            traceback.print_exc()
    print("\nHammasi tugadi.")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
