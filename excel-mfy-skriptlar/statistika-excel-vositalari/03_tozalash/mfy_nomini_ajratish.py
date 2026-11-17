"""
"Яшаш манзили" ustunidan MFY nomini ajratib, "МФЙ номи" ustuniga yozadi.

Kafolatlar:
  * Faqat "МФЙ номи" ustunidagi kataklar o'zgaradi. Boshqa ustunlar,
    varaqlar, shrift, rang, kenglik, filtr, birlashgan kataklar - o'z holicha.
  * Asl fayl o'zgarmaydi, natija yangi faylga yoziladi.
  * Saqlangandan so'ng skript o'zi asl fayl bilan solishtirib tekshiradi.

O'rnatish:  pip install openpyxl
Ishlatish:  python mfy_ajratish_v2.py [kirish.xlsx] [chiqish.xlsx]
"""
import csv
import os
import re
import sys
import zipfile
from openpyxl import load_workbook

# ===================== SOZLAMALAR =====================
INPUT_FILE = r"tayyor_Ижарадан_даромад_олганлар_реестри+++.xlsx"
OUTPUT_FILE = r"natija_MFY.xlsx"
REPORT_FILE = r"topilmagan_qatorlar.csv"
SHEETS = ["Мулк", "техника ва бошқалар"]   # свод ga tegilmaydi
KEEP_SUFFIX = True    # True -> "НАВРЎЗ МФЙ", False -> "НАВРЎЗ"
# Avtomatik topilmasa qo'lda bering, masalan:
# OVERRIDES = {"Мулк": {"target": "E", "address": "F"}}
OVERRIDES = {}
# ======================================================

if len(sys.argv) > 1:
    INPUT_FILE = sys.argv[1]
if len(sys.argv) > 2:
    OUTPUT_FILE = sys.argv[2]

# Nisbiy yo'llar skript turgan papkaga nisbatan olinadi
BASE = os.path.dirname(os.path.abspath(__file__))


def resolve(p):
    return p if os.path.isabs(p) else os.path.join(BASE, p)


INPUT_FILE, OUTPUT_FILE, REPORT_FILE = (
    resolve(INPUT_FILE), resolve(OUTPUT_FILE), resolve(REPORT_FILE)
)
if not os.path.exists(INPUT_FILE):
    print("Fayl topilmadi:", INPUT_FILE)
    print("Skript papkasidagi Excel fayllar:")
    for f in sorted(os.listdir(BASE)):
        if f.lower().endswith((".xlsx", ".xlsm")):
            print("   ", f)
    print("INPUT_FILE ga to'g'ri nomni (yoki to'liq yo'lni) yozing.")
    sys.exit(1)

SUFFIX = (
    r"(?:[МM][ФF][ЙYJ]"
    r"|[МM]СГ"
    r"|ШФЙ|SHFY"
    r"|маҳалласи|махалласи|mahallasi|mahalla)"
)
PREFIX = re.compile(r"^.*\b(?:тумани|tumani)\s+", re.I)
PATTERN = re.compile(
    rf"^(?P<name>.+?)\s+(?P<suf>{SUFFIX})(?:\s+Д\.?\s*\d.*)?$", re.I
)


def extract_mfy(address):
    """Manzildan MFY nomini qaytaradi, topilmasa None."""
    if address is None:
        return None
    for part in re.split(r"[,\n;]", str(address)):
        part = PREFIX.sub("", part.strip()).strip()
        m = PATTERN.match(part)
        if m:
            name = re.sub(r"\s+", " ", m.group("name")).strip(" .-")
            if not name:
                continue
            return f"{name} {m.group('suf')}" if KEEP_SUFFIX else name
    return None


def norm(v):
    return " ".join(str(v or "").lower().split())


def find_columns(ws):
    """(sarlavha_qatori, target_ustun, [manzil_ustunlari]) yoki None."""
    ov = OVERRIDES.get(ws.title)
    if ov:
        from openpyxl.utils import column_index_from_string as ci
        return 1, ci(ov["target"]), [ci(ov["address"])]

    for row in ws.iter_rows(min_row=1, max_row=6):
        exact = loose = None
        addr = []
        for c in row:
            v = norm(c.value)
            if not v:
                continue
            if v.replace(" ", "") in ("мфйноми", "mfynomi") and exact is None:
                exact = c.column
            elif (("мфй" in v or "mfy" in v) and "ном" in v
                  and "яшаш" not in v and "жойлаш" not in v
                  and "кўчмас" not in v and loose is None):
                loose = c.column
            if "яшаш" in v and "манзил" in v:
                addr.append(c.column)          # asosiy manzil ustuni
        # zaxira: "яшаш ... МФЙ номи" ustunida ham manzil bo'lishi mumkin
        for c in row:
            v = norm(c.value)
            if "яшаш" in v and c.column not in addr:
                addr.append(c.column)
        tgt = exact or loose
        if tgt and addr:
            addr = [a for a in addr if a != tgt]
            return row[0].row, tgt, addr
    return None


def warn_if_lossy(path):
    """openpyxl saqlay olmaydigan narsalar bormi?"""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    bad = [k for k in ("drawings/", "charts/", "pivotTables/", "media/",
                       "slicers/", "vbaProject", "externalLinks/")
           if any(k in n for n in names)]
    if bad:
        print("DIQQAT: faylda", bad, "bor - ular saqlanmasligi mumkin.")


def style_sig(c):
    """Katak stilining solishtiriladigan ko'rinishi."""
    return (repr(c.font), repr(c.fill), repr(c.alignment),
            repr(c.border), c.number_format)


def verify(src, dst, target_cols):
    """Asl va yangi fayl solishtiriladi: target ustundan tashqari hammasi bir xil."""
    a, b = load_workbook(src), load_workbook(dst)
    if a.sheetnames != b.sheetnames:
        return ["varaqlar ro'yxati farq qiladi"]
    problems = []
    for wa in a.worksheets:
        wb_ = b[wa.title]
        tcol = target_cols.get(wa.title)
        if (wa.max_row, wa.max_column) != (wb_.max_row, wb_.max_column):
            problems.append(f"{wa.title}: o'lcham farq qiladi")
        for row in wa.iter_rows():
            for ca in row:
                cb = wb_.cell(ca.row, ca.column)
                if ca.column != tcol and ca.value != cb.value:
                    problems.append(f"{wa.title}!{ca.coordinate}: qiymat o'zgargan")
                if style_sig(ca) != style_sig(cb):
                    problems.append(f"{wa.title}!{ca.coordinate}: format o'zgargan")
        for k, d in wa.column_dimensions.items():
            if d.width != wb_.column_dimensions[k].width:
                problems.append(f"{wa.title}: {k} ustun kengligi o'zgargan")
        if wa.merged_cells.ranges != wb_.merged_cells.ranges:
            problems.append(f"{wa.title}: birlashgan kataklar farq qiladi")
        if wa.auto_filter.ref != wb_.auto_filter.ref:
            problems.append(f"{wa.title}: filtr o'zgargan")
        if len(problems) > 20:
            break
    return problems


def main():
    warn_if_lossy(INPUT_FILE)
    wb = load_workbook(INPUT_FILE)
    wanted = {s.strip().lower() for s in SHEETS}
    target_cols, report = {}, []

    for ws in wb.worksheets:
        if ws.title.strip().lower() not in wanted:
            continue
        cols = find_columns(ws)
        if not cols:
            print(f"[{ws.title}] kerakli ustunlar topilmadi -> OVERRIDES bering")
            continue
        hdr, tgt, addr_cols = cols
        target_cols[ws.title] = tgt
        done = 0
        for r in range(hdr + 1, ws.max_row + 1):
            mfy = None
            for ac in addr_cols:
                mfy = extract_mfy(ws.cell(r, ac).value)
                if mfy:
                    break
            if mfy:
                cell = ws.cell(r, tgt)
                if cell.value != mfy:
                    cell.value = mfy          # faqat qiymat, stil o'zgarmaydi
                done += 1
            elif any(ws.cell(r, ac).value for ac in addr_cols):
                report.append([ws.title, r, ws.cell(r, addr_cols[0]).value])
        print(f"[{ws.title}] ustun: target={tgt}, manzil={addr_cols} | "
              f"yozildi: {done}")

    absent = wanted - {w.title.strip().lower() for w in wb.worksheets}
    if absent:
        print("Topilmagan varaqlar:", absent, "| Mavjud:", wb.sheetnames)

    wb.save(OUTPUT_FILE)
    print("Saqlandi:", OUTPUT_FILE)

    if report:
        with open(REPORT_FILE, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["varaq", "qator", "manzil"])
            w.writerows(report)
        print(f"MFY topilmagan qatorlar: {len(report)} ta -> {REPORT_FILE} "
              f"(Excelda o'sha katakka tegilmadi)")

    problems = verify(INPUT_FILE, OUTPUT_FILE, target_cols)
    if problems:
        print("TEKSHIRUV: MUAMMO topildi:")
        for p in problems[:20]:
            print("  -", p)
    else:
        print("TEKSHIRUV: OK - boshqa ustun, ma'lumot va formatlar bir xil.")


if __name__ == "__main__":
    main()
