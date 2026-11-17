"""
Bitta Excel faylning IKKI varag'idan (jismoniy shaxslar + yuridik shaxslar)
avval SURXONDARYO viloyatini ajratadi, keyin TUMAN bo'yicha alohida fayllarga
bo'ladi. Har bir tuman uchun bitta fayl: ichida ikki varaq ("Jismoniy shaxslar"
va "Yuridik shaxslar"), har varaq oxirida soni yoziladi.

Muammo: jismoniy varaqda tuman NOMI yo'q, faqat kod (district_code). Shuning
uchun skript "kod -> tuman nomi" jadvalini yuridik varaqdan (unda ham SOATO kodi,
ham TUMAN nomi bor) quradi va jismoniy varaqdagi kodlarni shu nom bilan bog'laydi.

Surxondaryo va uning kod prefiksi (1722) yuridik varaqdagi "HUDUD" nomidan
avtomatik aniqlanadi.

O'rnatish:  pip install openpyxl
Ishlatish:  python surxondaryo_tuman_bolish.py [fayl.xlsx]
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
INPUT_FILE = "Қўшимча_рўйхат_бизнесни_рўйхатга_олиш_учун_16_11_26.xlsx"
OUTPUT_DIR = None                 # None -> "<fayl nomi>_Surxondaryo_tumanlar"
REGION_NAME = r"sur\w*dar"        # viloyat nomida shu (Surxondaryo) qidiriladi
REGION_PREFIX = None              # None -> yuridik varaqdan aniqlanadi (odatda "1722")
COUNT_LABEL = "Jami:"
# Tuman nomi topilmagan kodlar uchun qo'lda nom berish (ixtiyoriy):
MANUAL_CODE_NAME = {}             # masalan {"1722202": "Angor tumani"}
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

IND_HINTS = {"fio": ["fio", "fish", "ifo", "f.i.o"],
             "region": ["region_code", "region"],
             "district": ["district_code", "district"]}
LEG_HINTS = {"hudud": ["hudud", "viloyat"], "tuman": ["tuman"],
             "soato": ["soato_code", "soato"]}


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def lat(v):
    s = nfc(v).casefold()
    return "".join(CYR.get(ch, ch) for ch in s)


def digits(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v))


def safe_filename(s):
    s = nfc(s).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    return re.sub(r"\s+", " ", s).strip(" .") or "tuman_yoq"


def match_cols(header, hints):
    cols = {}
    for j, v in enumerate(header):
        h = lat(v).strip()
        for role, words in hints.items():
            if role not in cols and any(w in h for w in words):
                cols[role] = j
    return cols


def read_header(ws):
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        return list(row)
    return []


def styled_header(ws, values, fill="1F4E78", color="FFFFFF"):
    out = []
    for v in values:
        c = WriteOnlyCell(ws, value=v)
        c.font = Font(bold=True, color=color)
        c.fill = PatternFill("solid", start_color=fill, end_color=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        out.append(c)
    return out


def count_row(ws, ncol, label):
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
    ind = leg = None     # (ws_title, header, cols)
    for ws in wb.worksheets:
        header = read_header(ws)
        if not header:
            continue
        ic = match_cols(header, IND_HINTS)
        lc = match_cols(header, LEG_HINTS)
        if "fio" in ic and ("district" in ic or "region" in ic) and ind is None:
            ind = (ws.title, header, ic)
        elif "tuman" in lc and ("hudud" in lc or "soato" in lc) and leg is None:
            leg = (ws.title, header, lc)
    if not leg:
        print("Yuridik (tuman nomi bor) varaq topilmadi.")
        return
    print(f"Jismoniy varaq: {ind[0] if ind else '— (topilmadi)'}")
    print(f"Yuridik varaq : {leg[0]}")

    # ---- 1) yuridik varaqdan: kod->tuman nomi, Surxondaryo prefiksi, va qatorlar
    code_name = defaultdict(Counter)
    prefix_cnt = Counter()
    leg_groups = defaultdict(list)     # tuman nomi -> [row,...]
    lc = leg[2]
    ws = wb[leg[0]]
    first = True
    for row in ws.iter_rows(values_only=True):
        if first:
            first = False
            continue
        row = list(row)
        soato = digits(row[lc["soato"]]) if "soato" in lc and lc["soato"] < len(row) else ""
        tuman = nfc(row[lc["tuman"]]).strip() if lc["tuman"] < len(row) else ""
        hud = row[lc["hudud"]] if "hudud" in lc and lc["hudud"] < len(row) else ""
        is_surxon = bool(REGION_RE.search(lat(hud)))
        if soato and tuman:
            code_name[soato[:7]][tuman] += 1
        if is_surxon and soato:
            prefix_cnt[soato[:4]] += 1
    prefix = REGION_PREFIX or (prefix_cnt.most_common(1)[0][0] if prefix_cnt else "1722")
    print(f"Surxondaryo kod prefiksi: {prefix}")

    code2name = {c: cnt.most_common(1)[0][0] for c, cnt in code_name.items()}

    # yuridik Surxondaryo qatorlarini tuman bo'yicha yig'amiz
    ws = wb[leg[0]]
    first = True
    for row in ws.iter_rows(values_only=True):
        if first:
            first = False
            continue
        row = list(row)
        soato = digits(row[lc["soato"]]) if "soato" in lc and lc["soato"] < len(row) else ""
        if not soato.startswith(prefix):
            continue
        tuman = nfc(row[lc["tuman"]]).strip() or f"Tuman {soato[:7]}"
        leg_groups[tuman].append(row)

    # ---- 2) jismoniy varaqdan: Surxondaryo qatorlari, tuman nomi kod orqali
    ind_groups = defaultdict(list)
    unmapped = Counter()
    if ind:
        ic = ind[2]
        ws = wb[ind[0]]
        first = True
        for row in ws.iter_rows(values_only=True):
            if first:
                first = False
                continue
            row = list(row)
            reg = digits(row[ic["region"]]) if "region" in ic and ic["region"] < len(row) else ""
            dist = digits(row[ic["district"]]) if "district" in ic and ic["district"] < len(row) else ""
            if not (reg.startswith(prefix) or dist.startswith(prefix)):
                continue
            d7 = dist[:7]
            name = MANUAL_CODE_NAME.get(d7) or code2name.get(d7)
            if not name:
                name = f"Nomalum tuman (kod {d7 or reg})"
                unmapped[d7 or reg] += 1
            ind_groups[name].append(row)
    wb.close()

    # ---- 3) har bir tuman uchun fayl
    out_dir = OUTPUT_DIR or os.path.join(
        BASE_DIR, os.path.splitext(os.path.basename(path))[0] + "_Surxondaryo_tumanlar")
    os.makedirs(out_dir, exist_ok=True)

    all_tumans = sorted(set(ind_groups) | set(leg_groups))
    summary = []
    for tuman in all_tumans:
        nb = Workbook(write_only=True)
        ic_n = len(ind[1]) if ind else 0
        lc_n = len(leg[1])
        a = len(ind_groups.get(tuman, []))      # YaTT soni
        b = len(leg_groups.get(tuman, []))      # Yuridik soni
        summ = f"{tuman} — YaTT: {a} ta, Yuridik: {b} ta"

        if ind and ind_groups.get(tuman):
            s1 = nb.create_sheet("YaTT")
            s1.append(styled_header(s1, ind[1], fill="375623"))
            for row in ind_groups[tuman]:
                s1.append(list(row) + [None] * (ic_n - len(row)))
            s1.append(count_row(s1, ic_n, summ))
            for j in range(ic_n):
                s1.column_dimensions[get_column_letter(j + 1)].width = 18
            s1.freeze_panes = "A2"

        if leg_groups.get(tuman):
            s2 = nb.create_sheet("Yuridik")
            s2.append(styled_header(s2, leg[1], fill="1F4E78"))
            for row in leg_groups[tuman]:
                s2.append(list(row) + [None] * (lc_n - len(row)))
            s2.append(count_row(s2, lc_n, summ))
            for j in range(lc_n):
                s2.column_dimensions[get_column_letter(j + 1)].width = 18
            s2.freeze_panes = "A2"

        if not nb.sheetnames:
            continue
        nb.save(os.path.join(out_dir, f"{safe_filename(tuman)}.xlsx"))
        ic_c = len(ind_groups.get(tuman, []))
        lc_c = len(leg_groups.get(tuman, []))
        summary.append((tuman, ic_c, lc_c))
        print(f"  {tuman:32} | jismoniy: {ic_c:5} | yuridik: {lc_c:5}")

    # umumiy ro'yxat
    ib = Workbook()
    iws = ib.active
    iws.title = "Jami"
    iws.append(["Tuman", "Jismoniy shaxslar", "Yuridik shaxslar", "Jami"])
    for c in iws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ti = tl = 0
    for name, a, b in sorted(summary, key=lambda x: -(x[1] + x[2])):
        iws.append([name, a, b, a + b])
        ti += a
        tl += b
        for c in iws[iws.max_row]:
            c.border = BOX
    iws.append(["JAMI", ti, tl, ti + tl])
    for c in iws[iws.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        c.border = BOX
    iws.column_dimensions["A"].width = 34
    for L in ("B", "C", "D"):
        iws.column_dimensions[L].width = 18
    iws.freeze_panes = "A2"
    iws.auto_filter.ref = f"A1:D{iws.max_row}"
    ib.save(os.path.join(out_dir, "00_Jami_royxat.xlsx"))

    print(f"Jami: {len(summary)} ta tuman | jismoniy {ti}, yuridik {tl}")
    if unmapped:
        print("DIQQAT: tuman nomi topilmagan kodlar (yuridik varaqda yo'q):")
        for code, c in unmapped.most_common(10):
            print(f"   kod {code}: {c} qator  (MANUAL_CODE_NAME ga nom qo'shing)")
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
