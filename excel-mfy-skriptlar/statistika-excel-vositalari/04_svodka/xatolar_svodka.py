"""
Yuklangan "xatolar" faylini TUMAN kesimida svodka qiladi: har tuman uchun
alohida chiroyli fayl (sarlavha, Jami qatori, foiz ustuni rangli) - Angor
namunasidagi ko'rinishda.

Fayl nomini yozish shart emas: skript o'z papkasidagi nomida "xatolar" bor
Excel faylni o'zi topadi. Natija "Tumanlar_svodka" papkasiga chiqadi.

O'rnatish (bir marta):  pip install openpyxl
Ishlatish:              xatolar_svodka.py ni ishga tushiring (yoki .bat ni bosing)
"""
import os
import re
import sys
import traceback
import unicodedata
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_KEYWORD = "xatolar"          # fayl nomida shu bo'lsa topadi
OUTPUT_DIR_NAME = "Tumanlar_svodka"
TITLE = ("Biznesni ro'yxatga olish tadbiri doirasida kelgan so'rovnomalarda "
         "yo'l qo'yilgan xato va kamchiliklar to'g'risida ma'lumot")
SHEET_NAME = "Mahallalar"
# ======================================================

CYR = {"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
       "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
       "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
       "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "",
       "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "ғ": "g",
       "қ": "q", "ҳ": "h"}
THIN = Side(style="thin", color="B0B0B0")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# chiqish ustunlari: (sarlavha, manba kaliti)
OUT_COLS = [
    ("№", None),
    ("Tuman", "tuman"),
    ("STIR", "stir"),
    ("Mahalla", "mahalla"),
    ("Jami xatolar", "jami"),
    ("Qayta yuborilgan", "qayta"),
    ("To'g'ri tuzatilgan", "togri"),
    ("Takroriy xato", "takroriy"),
    ("Tekshirilmoqda", "tekshir"),
    ("Qolgan xato soni", "qolgan"),
    ("Tuzatilgan ulushi", "ulush"),
]
NUM_KEYS = ["jami", "qayta", "togri", "takroriy", "tekshir", "qolgan"]
HINTS = {
    "viloyat": ["viloyat", "вилоят", "hudud"],
    "tuman": ["tuman", "туман"],
    "stir": ["stir", "стир", "inn"],
    "mahalla": ["mahalla", "маҳалла", "махалла"],
    "jami": ["jami xato", "жами хато", "jami xatolar"],
    "qayta": ["qayta yub", "қайта юб", "qayta yuborilgan"],
    "togri": ["to'g'ri tuz", "togri tuz", "тўғри туз", "to‘g‘ri"],
    "takroriy": ["takroriy", "такрорий"],
    "tekshir": ["tekshir", "текшир"],
    "qolgan": ["kutilmoqda", "кутилмоқда", "qolgan"],
    "ulush": ["ulush", "улуш"],
}


def nfc(v):
    return "" if v is None else unicodedata.normalize("NFC", str(v))


def lat(v):
    return "".join(CYR.get(ch, ch) for ch in nfc(v).casefold())


def num(v):
    if v is None or nfc(v).strip() == "":
        return 0
    if isinstance(v, (int, float)):
        return v
    s = re.sub(r"[^\d.\-]", "", str(v).replace(",", "."))
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        return 0


def safe_filename(s):
    s = nfc(s).strip()
    s = re.sub(r'[\\/:*?"<>|]', "_", s)
    return re.sub(r"\s+", " ", s).strip(" .") or "tuman_yoq"


def find_input():
    files = [f for f in sorted(os.listdir(BASE_DIR))
             if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    hits = [f for f in files if INPUT_KEYWORD in lat(f)]
    cand = hits if hits else files
    if len(cand) == 1:
        return os.path.join(BASE_DIR, cand[0])
    print("Qaysi fayl?" if hits else "'xatolar' fayli topilmadi. Fayllar:")
    for i, f in enumerate(cand, 1):
        print(f"  {i}. {f}")
    try:
        return os.path.join(BASE_DIR, cand[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def detect(ws):
    """Sarlavha qatori va ustun indekslarini topadi."""
    for r in range(1, min(ws.max_row, 20) + 1):
        vals = [c.value for c in ws[r]]
        cols = {}
        for j, v in enumerate(vals):
            h = lat(v).strip()
            if not h:
                continue
            for role, words in HINTS.items():
                if role not in cols and any(w in h for w in words):
                    cols[role] = j
        if "mahalla" in cols and "jami" in cols and "tuman" in cols:
            return r, cols
    return None, None


def main():
    src = find_input()
    if not src:
        print("Fayl tanlanmadi.")
        return
    print("Fayl:", os.path.basename(src))
    wb = load_workbook(src, read_only=True)
    ws = None
    for sh in wb.worksheets:
        hdr, cols = detect(sh)
        if hdr:
            ws, hdr_row, ci = sh, hdr, cols
            break
    if ws is None:
        print("Sarlavha (Mahalla/Tuman/Jami xatolar) topilmadi.")
        return
    print(f"Varaq: [{ws.title}] | sarlavha qatori: {hdr_row}")

    get = lambda row, k: row[ci[k]] if k in ci and ci[k] < len(row) else None
    groups = defaultdict(list)
    rows = ws.iter_rows(min_row=hdr_row + 1, values_only=True)
    for row in rows:
        tuman = nfc(get(row, "tuman")).strip()
        mahalla = nfc(get(row, "mahalla")).strip()
        # "Jami" (yig'indi) qatorini va bo'sh qatorlarni tashlaymiz
        if not tuman or not mahalla or lat(mahalla) == "jami" or lat(tuman) == "jami":
            continue
        rec = {
            "tuman": tuman, "stir": nfc(get(row, "stir")).strip(), "mahalla": mahalla,
            "jami": num(get(row, "jami")), "qayta": num(get(row, "qayta")),
            "togri": num(get(row, "togri")), "takroriy": num(get(row, "takroriy")),
            "tekshir": num(get(row, "tekshir")), "qolgan": num(get(row, "qolgan")),
        }
        groups[tuman].append(rec)
    wb.close()
    if not groups:
        print("Ma'lumot topilmadi.")
        return

    out_dir = os.path.join(BASE_DIR, OUTPUT_DIR_NAME)
    os.makedirs(out_dir, exist_ok=True)
    today = datetime.now().strftime("%d.%m.%Y %H:%M")
    summary = []

    for tuman in sorted(groups, key=lambda t: t.lower()):
        recs = groups[tuman]
        build_sheet(out_dir, tuman, recs, today)
        tj = sum(r["jami"] for r in recs)
        tq = sum(r["qayta"] for r in recs)
        summary.append((tuman, len(recs), tj, tq, tq / tj if tj else 0))
        print(f"  {tuman:26} | {len(recs):4} mahalla | jami xato {tj}")

    build_summary(out_dir, summary, today)
    print(f"Jami: {len(summary)} ta tuman")
    print("Papka:", out_dir)


def build_sheet(out_dir, tuman, recs, today):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ncol = len(OUT_COLS)
    last_L = get_column_letter(ncol)

    # sarlavha
    ws.merge_cells(f"A1:{last_L}3")
    t = ws.cell(1, 1)
    t.value = TITLE
    t.font = Font(name="Times New Roman", size=13, bold=True)
    t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(4, ncol).value = today
    ws.cell(4, ncol).alignment = Alignment(horizontal="right")
    ws.cell(4, ncol).font = Font(italic=True, size=9, color="595959")

    hdr_row = 5
    for j, (title, _) in enumerate(OUT_COLS, 1):
        c = ws.cell(hdr_row, j)
        c.value = title
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F6E43", end_color="1F6E43")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[hdr_row].height = 34

    # Jami (yig'indi) qatori
    tot = {k: sum(r[k] for r in recs) for k in NUM_KEYS}
    jr = hdr_row + 1
    ws.cell(jr, 4).value = "JAMI"
    for j, (_, key) in enumerate(OUT_COLS, 1):
        c = ws.cell(jr, j)
        if key in NUM_KEYS:
            c.value = tot[key]
        elif key == "ulush":
            c.value = (tot["qayta"] / tot["jami"]) if tot["jami"] else 0
            c.number_format = "0%"
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        c.border = BOX
        c.alignment = Alignment(horizontal="center" if j != 4 else "left",
                                vertical="center")

    # ma'lumot qatorlari
    r = jr + 1
    for i, rec in enumerate(sorted(recs, key=lambda x: x["mahalla"].lower()), 1):
        ulush = (rec["qayta"] / rec["jami"]) if rec["jami"] else 0
        values = {"№": i, "tuman": rec["tuman"], "stir": rec["stir"],
                  "mahalla": rec["mahalla"], "ulush": ulush, **{k: rec[k] for k in NUM_KEYS}}
        for j, (title, key) in enumerate(OUT_COLS, 1):
            c = ws.cell(r, j)
            c.value = values["№"] if title == "№" else values.get(key)
            c.border = BOX
            if key == "ulush":
                c.number_format = "0%"
                c.alignment = Alignment(horizontal="center")
            elif key in NUM_KEYS or title == "№":
                c.alignment = Alignment(horizontal="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
        if i % 2 == 0:
            for j in range(1, ncol + 1):
                if ws.cell(r, j).fill.fill_type != "solid":
                    ws.cell(r, j).fill = PatternFill(
                        "solid", start_color="F3F7F3", end_color="F3F7F3")
        r += 1

    # foiz ustuniga rang shkalasi (qizil -> sariq -> yashil)
    ulush_col = get_column_letter(ncol)
    rng = f"{ulush_col}{jr}:{ulush_col}{r - 1}"
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type="num", start_value=0, start_color="F8696B",
        mid_type="num", mid_value=0.5, mid_color="FFEB84",
        end_type="num", end_value=1, end_color="63BE7B"))

    # kengliklar, muzlatish, filtr, chop
    widths = [5, 18, 12, 26, 11, 13, 13, 11, 13, 12, 12]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f"A{jr + 1}"
    ws.auto_filter.ref = f"A{hdr_row}:{last_L}{r - 1}"
    ws.print_title_rows = f"{hdr_row}:{hdr_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False

    out = os.path.join(out_dir, safe_filename(tuman) + ".xlsx")
    wb.save(out)


def build_summary(out_dir, summary, today):
    wb = Workbook()
    ws = wb.active
    ws.title = "Svodka"
    head = ["№", "Tuman", "Mahalla soni", "Jami xatolar", "Qayta yuborilgan",
            "Tuzatilgan ulushi"]
    ws.merge_cells("A1:F1")
    ws.cell(1, 1).value = "Tumanlar kesimida xatolar svodkasi"
    ws.cell(1, 1).font = Font(size=14, bold=True, color="1F6E43")
    ws.cell(1, 1).alignment = Alignment(horizontal="center")
    ws.cell(2, 6).value = today
    ws.cell(2, 6).font = Font(italic=True, size=9, color="595959")
    ws.append(head)
    hr = 3
    for c in ws[hr]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F6E43", end_color="1F6E43")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    tj = tq = 0
    for i, (t, nm, j, q, u) in enumerate(sorted(summary, key=lambda x: -x[2]), 1):
        ws.append([i, t, nm, j, q, u])
        ws.cell(ws.max_row, 6).number_format = "0%"
        for c in ws[ws.max_row]:
            c.border = BOX
            c.alignment = Alignment(horizontal="left" if c.column == 2 else "center")
        tj += j
        tq += q
    ws.append(["", "JAMI", sum(s[1] for s in summary), tj, tq, (tq / tj if tj else 0)])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        c.border = BOX
        if c.column == 6:
            c.number_format = "0%"
    ws.conditional_formatting.add(
        f"F{hr + 1}:F{ws.max_row}", ColorScaleRule(
            start_type="num", start_value=0, start_color="F8696B",
            mid_type="num", mid_value=0.5, mid_color="FFEB84",
            end_type="num", end_value=1, end_color="63BE7B"))
    for j, w in enumerate([5, 24, 14, 14, 16, 16], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = f"A{hr + 1}"
    ws.auto_filter.ref = f"A{hr}:F{ws.max_row}"
    ws.sheet_view.showGridLines = False
    wb.save(os.path.join(out_dir, "00_Jami_svodka.xlsx"))


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
