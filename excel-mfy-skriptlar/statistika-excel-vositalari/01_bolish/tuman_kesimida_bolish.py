"""
Bir faylning HAR BIR varag'ini (ikkala listni ham) TUMAN bo'yicha alohida
fayllarga bo'ladi. Har bir tuman uchun bitta fayl: ichida har bir manba varaq
(masalan "Сўровномалар" va "Интервьюерлар") o'z nomi bilan alohida varaq bo'ladi,
oxirida soni yoziladi.

Har varaqdagi TUMAN ustunini o'zi topadi (viloyat/hudud ustuni bilan aralashmaydi).
Agar viloyat ustuni bo'lsa, faqat Surxondaryo qatorlari olinadi (REGION_NAME).

O'rnatish:  pip install openpyxl
Ishlatish:  python tuman_kesimida_bolish.py [fayl.xlsx]
"""
import os
import re
import sys
import time
import traceback
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = "Ҳоким_ёрдамчилари_саволномаларининг_хатоликлари_16_11_26+.xlsx"
OUTPUT_DIR = None                  # None -> "<fayl nomi>_tumanlar"
REGION_NAME = r"sur\w*dar"         # viloyat ustuni bo'lsa, faqat shu (Surxondaryo)
ONLY_REGION = True                 # False -> viloyatdan qat'i nazar hamma tuman
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
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def lat(v):
    return "".join(CYR.get(ch, ch) for ch in nfc(v).casefold())


def key(v):
    return " ".join(lat(v).split())


def safe_filename(s):
    s = nfc(s).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    return re.sub(r"\s+", " ", s).strip(" .") or "tuman_yoq"


def read_header(ws):
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return list(row)
    return []


def find_cols(header):
    """-> (tuman_ustun_idx, viloyat_ustun_idx|None)"""
    tcol = vcol = None
    for j, v in enumerate(header):
        h = lat(v).strip()
        if not h:
            continue
        if vcol is None and ("hudud" in h or "viloyat" in h or "вилоят" in lat(v)):
            vcol = j
        if tcol is None and ("tuman" in h or "туман" in lat(v)) \
                and "viloyat" not in h and "hudud" not in h:
            tcol = j
    return tcol, vcol


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

    wb = load_workbook(path, read_only=True)
    sheets = []          # (title, header, tcol, vcol)
    for ws in wb.worksheets:
        header = read_header(ws)
        if not header:
            continue
        tcol, vcol = find_cols(header)
        if tcol is None:
            print(f"  [{ws.title}] 'Туман' ustuni topilmadi - o'tkazildi")
            continue
        sheets.append((ws.title, header, tcol, vcol))
        print(f"  [{ws.title}] Туман={get_column_letter(tcol + 1)}"
              + (f", viloyat={get_column_letter(vcol + 1)}" if vcol is not None else ""))
    if not sheets:
        print("Hech bir varaqda 'Туман' ustuni topilmadi.")
        return

    # har varaqni tuman bo'yicha yig'amiz
    data = {}            # sheet_title -> {tuman_key: [rows]}
    disp = {}            # tuman_key -> ko'rsatiladigan nom
    ncols = {}
    skipped = Counter()
    for title, header, tcol, vcol in sheets:
        ws = wb[title]
        ncols[title] = len(header)
        groups = defaultdict(list)
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            row = list(row)
            if ONLY_REGION and vcol is not None:
                v = row[vcol] if vcol < len(row) else None
                if not REGION_RE.search(lat(v)):
                    skipped[title] += 1
                    continue
            tv = row[tcol] if tcol < len(row) else None
            k = key(tv)
            if not k:
                k = "__bosh__"
                disp.setdefault(k, "Tuman ko'rsatilmagan")
            else:
                disp.setdefault(k, nfc(tv).strip())
            groups[k].append(row)
        data[title] = groups
    wb.close()

    all_keys = sorted(set().union(*[set(g) for g in data.values()]),
                      key=lambda k: disp.get(k, k).lower())
    out_dir = OUTPUT_DIR or os.path.join(
        BASE_DIR, os.path.splitext(os.path.basename(path))[0] + "_tumanlar")
    os.makedirs(out_dir, exist_ok=True)

    palette = ["1F4E78", "375623", "833C00", "7030A0"]
    summary = []
    for k in all_keys:
        name = disp.get(k, k)
        nb = Workbook(write_only=True)
        counts = []
        for si, (title, header, tcol, vcol) in enumerate(sheets):
            rows = data[title].get(k, [])
            counts.append(len(rows))
            if not rows:
                continue
            ws = nb.create_sheet(title[:31])
            nc = ncols[title]
            ws.append(styled_header(ws, header, palette[si % len(palette)]))
            for row in rows:
                ws.append(list(row) + [None] * (nc - len(row)))
            ws.append(count_cells(ws, nc, f"{COUNT_LABEL} {len(rows)} ta"))
            for j in range(nc):
                ws.column_dimensions[get_column_letter(j + 1)].width = 18
            ws.freeze_panes = "A2"
        if not nb.sheetnames:
            continue
        nb.save(os.path.join(out_dir, f"{safe_filename(name)}.xlsx"))
        summary.append((name, counts))
        print(f"  {name:28} | " + " | ".join(
            f"{sheets[i][0][:14]}: {c}" for i, c in enumerate(counts)))

    # umumiy ro'yxat
    ib = Workbook()
    iws = ib.active
    iws.title = "Jami"
    head = ["Tuman"] + [t for t, *_ in sheets] + ["Jami"]
    iws.append(head)
    for c in iws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    totals = [0] * len(sheets)
    for name, counts in sorted(summary, key=lambda x: -sum(x[1])):
        iws.append([name] + counts + [sum(counts)])
        for i, c in enumerate(counts):
            totals[i] += c
        for c in iws[iws.max_row]:
            c.border = BOX
    iws.append(["JAMI"] + totals + [sum(totals)])
    for c in iws[iws.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        c.border = BOX
    iws.column_dimensions["A"].width = 30
    for j in range(2, len(head) + 1):
        iws.column_dimensions[get_column_letter(j)].width = 18
    iws.freeze_panes = "A2"
    iws.auto_filter.ref = f"A1:{get_column_letter(len(head))}{iws.max_row}"
    ib.save(os.path.join(out_dir, "00_Jami_royxat.xlsx"))

    print(f"Jami: {len(summary)} ta tuman")
    for i, (t, *_) in enumerate(sheets):
        print(f"  {t}: {totals[i]} ta" + (f" (boshqa viloyat o'tkazildi: {skipped[t]})"
                                          if skipped[t] else ""))
    print(f"Vaqt: {time.time() - t0:.1f} s")
    print("Papka:", out_dir)
    print("Umumiy ro'yxat: 00_Jami_royxat.xlsx")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
