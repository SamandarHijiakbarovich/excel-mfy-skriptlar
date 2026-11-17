"""
YaTT o'tmaganlar faylini TUMAN papka -> MAHALLA fayllarga bo'ladi.

Papka ko'rinishi:
  YaTT/
      Tumanlar/
          <Tuman nomi>/
              <Mahalla>_Qo'shimcha_yatt.xlsx
      00_Jami_royxat.xlsx

Sarlavha qatori va Tuman/Mahalla (MFY) ustunini o'zi topadi. Viloyat ustuni
bo'lsa, faqat Surxondaryo olinadi.

O'rnatish:  pip install openpyxl
Ishlatish:  python yatt_mahalla_bolish.py [fayl.xlsx]
"""
import os
import re
import shutil
import sys
import time
import traceback
import unicodedata
from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_KEYWORD = "yatt"       # fayl nomida shu bo'lsa o'zi topadi
ROOT_FOLDER = "YaTT"
FILE_SUFFIX = "_Qo'shimcha_yatt"           # mahalla fayl nomi oxiri
REGION_NAME = r"sur\w*dar"
ONLY_REGION = True
STRIP_MFY = True
COUNT_LABEL = "Jami:"
# ======================================================

if len(sys.argv) > 1:
    INPUT_FILE_ARG = sys.argv[1]
else:
    INPUT_FILE_ARG = None

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
EMPTY = {"", "0", "00", "#н/д", "#n/a", "n/a", "nan", "none", "-", "—", "#н/а"}
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def lat(v):
    return "".join(CYR.get(ch, ch) for ch in nfc(v).casefold())


def is_empty(v):
    return nfc(v).strip().casefold() in EMPTY


def clean_mah(v):
    s = QUOTES_RE.sub("", nfc(v).strip())
    s = MFY_TOKEN.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip(" -.,")


def norm_key(v):
    return re.sub(r"[^a-z0-9]", "", lat(clean_mah(v)))


def disp_mah(v):
    s = clean_mah(v) if STRIP_MFY else nfc(v).strip()
    return s or nfc(v).strip()


def safe_filename(s):
    s = nfc(s).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    return re.sub(r"\s+", " ", s).strip(" .") or "nomsiz"


def find_header(ws):
    """(sarlavha_qatori, tuman_idx, mahalla_idx, viloyat_idx) yoki None."""
    for r in range(1, 8):
        vals = None
        for row in ws.iter_rows(min_row=r, max_row=r, values_only=True):
            vals = list(row)
        if not vals:
            continue
        tcol = mcol = vcol = None
        for j, v in enumerate(vals):
            h = lat(v).strip()
            if not h:
                continue
            if vcol is None and ("viloyat" in h or "hudud" in h):
                vcol = j
            if (tcol is None and "tuman" in h and "viloyat" not in h
                    and "hudud" not in h and "soato" not in h):
                tcol = j
            if mcol is None and ("mahalla" in h or "maxalla" in h or "mfy" in h):
                mcol = j
        if tcol is not None and mcol is not None:
            return r, tcol, mcol, vcol
    return None


def styled_header(ws, values):
    out = []
    for v in values:
        c = WriteOnlyCell(ws, value=v)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
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
        cells.append(c)
    return cells


def find_input():
    if INPUT_FILE_ARG:
        p = INPUT_FILE_ARG if os.path.isabs(INPUT_FILE_ARG) else os.path.join(BASE_DIR, INPUT_FILE_ARG)
        return p if os.path.exists(p) else None
    files = [f for f in sorted(os.listdir(BASE_DIR))
             if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    hits = [f for f in files if INPUT_KEYWORD in lat(f)]
    if len(hits) == 1:
        return os.path.join(BASE_DIR, hits[0])
    cand = hits or files
    print("Qaysi fayl?" if hits else f"'{INPUT_KEYWORD}' bor fayl topilmadi. Fayllar:")
    for i, f in enumerate(cand, 1):
        print(f"  {i}. {f}")
    try:
        return os.path.join(BASE_DIR, cand[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def main():
    src = find_input()
    if not src:
        print("Fayl topilmadi.")
        return
    print("Fayl:", os.path.basename(src))
    t0 = time.time()

    wb = load_workbook(src, read_only=True)
    found = None
    for sh in wb.worksheets:
        fh = find_header(sh)
        if fh:
            found = (sh.title,) + fh
            break
    if not found:
        print("Tuman va Mahalla(MFY) ustunlari topilmadi.")
        return
    title, hdr_row, tcol, mcol, vcol = found
    ws = wb[title]
    header = None
    for row in ws.iter_rows(min_row=hdr_row, max_row=hdr_row, values_only=True):
        header = list(row)
    ncol = len(header)
    print(f"Varaq: [{title}] | sarlavha: {hdr_row}-qator | "
          f"Tuman={get_column_letter(tcol + 1)}, Mahalla={get_column_letter(mcol + 1)}")

    tuman_disp, mah_disp = {}, {}
    data = defaultdict(lambda: defaultdict(list))   # tuman -> mahalla -> rows
    skipped = 0
    for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
        row = list(row)
        if all(is_empty(x) for x in row):
            continue
        if ONLY_REGION and vcol is not None:
            if not REGION_RE.search(lat(row[vcol] if vcol < len(row) else "")):
                skipped += 1
                continue
        tv = row[tcol] if tcol < len(row) else None
        mv = row[mcol] if mcol < len(row) else None
        if is_empty(tv):
            tk, tdisp = "__tuman_yoq__", "Tuman korsatilmagan"
        else:
            tk, tdisp = " ".join(lat(tv).split()), nfc(tv).strip()
        if is_empty(mv):
            mk, mdisp = "__mahalla_yoq__", "Mahalla korsatilmagan"
        else:
            mk = norm_key(mv) or "__mahalla_yoq__"
            mdisp = disp_mah(mv) or "Mahalla korsatilmagan"
        tuman_disp.setdefault(tk, tdisp)
        mah_disp.setdefault((tk, mk), mdisp)
        data[tk][mk].append(row)
    wb.close()
    if not data:
        print("Ma'lumot topilmadi.")
        return

    root = os.path.join(BASE_DIR, ROOT_FOLDER)
    if os.path.isdir(root):
        try:
            shutil.rmtree(root)
            print("Eski papka tozalandi.")
        except Exception as e:
            print("Eski papkani tozalab bo'lmadi:", e)
    tumanlar_dir = os.path.join(root, "Tumanlar")
    os.makedirs(tumanlar_dir, exist_ok=True)

    summary = []
    n_files = 0
    for tk in sorted(data, key=lambda k: tuman_disp[k].lower()):
        tdir = os.path.join(tumanlar_dir, safe_filename(tuman_disp[tk]))
        os.makedirs(tdir, exist_ok=True)
        mah_total = 0
        for mk in sorted(data[tk], key=lambda k: mah_disp[(tk, k)].lower()):
            rows = data[tk][mk]
            name = mah_disp[(tk, mk)]
            nb = Workbook(write_only=True)
            ws2 = nb.create_sheet(title[:31] or "Ro'yxat")
            ws2.append(styled_header(ws2, header))
            for row in rows:
                ws2.append(list(row) + [None] * (ncol - len(row)))
            ws2.append(count_cells(ws2, ncol, f"{COUNT_LABEL} {len(rows)} ta"))
            for j in range(ncol):
                ws2.column_dimensions[get_column_letter(j + 1)].width = 20
            ws2.freeze_panes = "A2"
            fp = os.path.join(tdir, safe_filename(name) + FILE_SUFFIX + ".xlsx")
            k = 2
            while os.path.exists(fp):
                fp = os.path.join(tdir, f"{safe_filename(name)} ({k}){FILE_SUFFIX}.xlsx")
                k += 1
            nb.save(fp)
            n_files += 1
            mah_total += 1
            summary.append((tuman_disp[tk], name, len(rows)))
        print(f"  {tuman_disp[tk]:26} | {mah_total} ta mahalla")

    # umumiy ro'yxat
    ib = Workbook()
    iws = ib.active
    iws.title = "Jami"
    iws.append(["Tuman", "Mahalla", "Qatorlar soni"])
    for c in iws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    total = 0
    for t, m, n in summary:
        iws.append([t, m, n])
        total += n
        for c in iws[iws.max_row]:
            c.border = BOX
    iws.append(["JAMI", f"{len(summary)} ta mahalla", total])
    for c in iws[iws.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        c.border = BOX
    iws.column_dimensions["A"].width = 24
    iws.column_dimensions["B"].width = 30
    iws.column_dimensions["C"].width = 16
    iws.freeze_panes = "A2"
    iws.auto_filter.ref = f"A1:C{iws.max_row}"
    ib.save(os.path.join(root, "00_Jami_royxat.xlsx"))

    print(f"Jami: {len(data)} ta tuman, {n_files} ta mahalla fayli"
          + (f" (boshqa viloyat o'tkazildi: {skipped})" if skipped else ""))
    print(f"Vaqt: {time.time() - t0:.1f} s")
    print("Papka:", root)


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
    if sys.stdin and sys.stdin.isatty():
        input("\nTugatish uchun Enter bosing...")
