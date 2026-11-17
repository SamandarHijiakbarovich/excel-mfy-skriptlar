"""
SVODKA: Surxondaryo tumanlari kesimida YaTT (jismoniy) va Yuridik shaxslar soni.
Bitta faylning ikki varag'idan (jismoniy + yuridik) sanab, bitta chiroyli jadval
chiqaradi. Faylni bo'lmaydi - faqat svodka.

Jismoniy varaqda tuman nomi yo'q (faqat kod), shuning uchun "kod -> tuman nomi"
jadvali yuridik varaqdan quriladi. Surxondaryo (1722) avtomatik aniqlanadi.

O'rnatish:  pip install openpyxl
Ishlatish:  python svodka_yatt_yuridik.py [fayl.xlsx]
"""
import os
import re
import sys
import traceback
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = "Қўшимча_рўйхат_бизнесни_рўйхатга_олиш_учун_16_11_26.xlsx"
OUTPUT_NAME = "Svodka_YaTT_Yuridik.xlsx"
REGION_NAME = r"sur\w*dar"
REGION_PREFIX = None               # None -> avtomatik (1722)
SORT_BY = "name"                   # "name" (alifbo) yoki "total" (jami bo'yicha)
MANUAL_CODE_NAME = {}              # {"1722202": "Angor tumani"}
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
THIN = Side(style="thin", color="9DB2CE")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
IND_HINTS = {"fio": ["fio", "fish", "ifo", "f.i.o"],
             "region": ["region_code", "region"],
             "district": ["district_code", "district"]}
LEG_HINTS = {"hudud": ["hudud", "viloyat"], "tuman": ["tuman"],
             "soato": ["soato_code", "soato"]}


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def lat(v):
    return "".join(CYR.get(ch, ch) for ch in nfc(v).casefold())


def digits(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v))


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


def resolve(name):
    p = name if os.path.isabs(name) else os.path.join(BASE_DIR, name)
    if os.path.exists(p):
        return p
    for f in sorted(os.listdir(BASE_DIR)):
        if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$") \
                and re.sub(r"[^a-z0-9]", "", f.lower()) == re.sub(r"[^a-z0-9]", "", name.lower()):
            return os.path.join(BASE_DIR, f)
    return p


def main():
    path = resolve(INPUT_FILE)
    if not os.path.exists(path):
        print("Fayl topilmadi:", path, "\nPapkadagi Excel fayllar:")
        for f in sorted(os.listdir(BASE_DIR)):
            if f.lower().endswith((".xlsx", ".xlsm")):
                print("  ", f)
        return
    print("Fayl:", os.path.basename(path))
    wb = load_workbook(path, read_only=True)
    ind = leg = None
    for ws in wb.worksheets:
        header = read_header(ws)
        if not header:
            continue
        ic, lc = match_cols(header, IND_HINTS), match_cols(header, LEG_HINTS)
        if "fio" in ic and ("district" in ic or "region" in ic) and ind is None:
            ind = (ws.title, ic)
        elif "tuman" in lc and ("hudud" in lc or "soato" in lc) and leg is None:
            leg = (ws.title, lc)
    if not leg:
        print("Yuridik (tuman nomi bor) varaq topilmadi.")
        return

    # kod -> tuman nomi + Surxondaryo prefiksi + yuridik sanog'i
    lc = leg[1]
    code_name = defaultdict(Counter)
    prefix_cnt = Counter()
    yuridik = Counter()
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
        if soato and tuman:
            code_name[soato[:7]][tuman] += 1
        if REGION_RE.search(lat(hud)) and soato:
            prefix_cnt[soato[:4]] += 1
    prefix = REGION_PREFIX or (prefix_cnt.most_common(1)[0][0] if prefix_cnt else "1722")
    code2name = {c: cnt.most_common(1)[0][0] for c, cnt in code_name.items()}

    ws = wb[leg[0]]
    first = True
    for row in ws.iter_rows(values_only=True):
        if first:
            first = False
            continue
        row = list(row)
        soato = digits(row[lc["soato"]]) if "soato" in lc and lc["soato"] < len(row) else ""
        if soato.startswith(prefix):
            yuridik[nfc(row[lc["tuman"]]).strip() or f"Tuman {soato[:7]}"] += 1

    # YaTT sanog'i
    yatt = Counter()
    if ind:
        ic = ind[1]
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
            name = MANUAL_CODE_NAME.get(d7) or code2name.get(d7) or f"Nomalum tuman (kod {d7 or reg})"
            yatt[name] += 1
    wb.close()

    tumans = set(yatt) | set(yuridik)
    rows = [(t, yatt[t], yuridik[t]) for t in tumans]
    if SORT_BY == "total":
        rows.sort(key=lambda x: -(x[1] + x[2]))
    else:
        rows.sort(key=lambda x: nfc(x[0]).lower())

    # ---- svodka jadvali
    nb = Workbook()
    ws = nb.active
    ws.title = "Svodka"
    ncol = 5
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    tc = ws.cell(1, 1)
    tc.value = ("Surxondaryo viloyati — tumanlar kesimida YaTT va yuridik "
                "shaxslar soni")
    tc.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.cell(2, 1).value = f"Sana: {datetime.now():%d.%m.%Y %H:%M}"
    ws.cell(2, 1).font = Font(size=10, italic=True, color="595959")

    head = ["№", "Tuman", "YaTT (jismoniy)", "Yuridik shaxslar", "Jami"]
    hr = 3
    for j, h in enumerate(head, 1):
        c = ws.cell(hr, j)
        c.value = h
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[hr].height = 30

    r = hr + 1
    for i, (t, a, b) in enumerate(rows, 1):
        ws.cell(r, 1).value = i
        ws.cell(r, 2).value = t
        ws.cell(r, 3).value = a
        ws.cell(r, 4).value = b
        ws.cell(r, 5).value = a + b
        ws.cell(r, 5).font = Font(bold=True)
        for j in range(1, ncol + 1):
            cell = ws.cell(r, j)
            cell.border = BOX
            cell.alignment = Alignment(horizontal="left" if j == 2 else "center",
                                       vertical="center")
            if j >= 3:
                cell.number_format = "#,##0"
        if i % 2 == 0:
            for j in range(1, ncol + 1):
                ws.cell(r, j).fill = PatternFill("solid", start_color="EAF0F8",
                                                 end_color="EAF0F8")
        r += 1

    # JAMI qatori
    ta, tb = sum(a for _, a, _ in rows), sum(b for _, _, b in rows)
    ws.cell(r, 1).value = ""
    ws.cell(r, 2).value = "JAMI"
    ws.cell(r, 3).value = ta
    ws.cell(r, 4).value = tb
    ws.cell(r, 5).value = ta + tb
    for j in range(1, ncol + 1):
        c = ws.cell(r, j)
        c.font = Font(bold=True, size=11)
        c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        c.border = BOX
        c.alignment = Alignment(horizontal="left" if j == 2 else "center")
        if j >= 3:
            c.number_format = "#,##0"

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 32
    for L in ("C", "D", "E"):
        ws.column_dimensions[L].width = 18
    ws.freeze_panes = f"A{hr + 1}"
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(ncol)}{r - 1}"
    ws.print_title_rows = f"{hr}:{hr}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False

    out = os.path.join(BASE_DIR, OUTPUT_NAME)
    try:
        nb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"Svodka_{datetime.now():%H%M%S}.xlsx")
        print("Svodka fayli ochiq ekan, boshqa nomga saqlanmoqda.")
        nb.save(out)

    print(f"Surxondaryo prefiksi: {prefix}")
    print(f"Tumanlar: {len(rows)} | YaTT jami: {ta} | Yuridik jami: {tb} | "
          f"Umumiy: {ta + tb}")
    for t, a, b in rows:
        print(f"  {t:32} | YaTT {a:6} | Yuridik {b:6} | Jami {a + b:6}")
    print("Saqlandi:", out)


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
