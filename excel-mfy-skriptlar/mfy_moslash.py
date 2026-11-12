# -*- coding: utf-8 -*-
"""
Asosiy fayldagi mahalla nomlarini "МФЙ лар руйхати.xlsx" dagi TO'G'RI nomlar bilan
TUMAN bo'yicha solishtiradi va yoniga qo'yadi (moslik foizi bilan).

Asosiy faylga "Mahalla nomi" ustunidan keyin (yo'q bo'lsa "Mahalla"dan keyin) 4 ta ustun qo'shiladi:
  MFY nomi (ro'yxat)  — ro'yxatdagi to'g'ri nom
  Moslik %            — 100 = aynan bir xil; 90% dan past bo'lsa nom qo'yilmaydi
  MFY STIR            — ro'yxatdagi СТИР
  Ro'yxat tumani      — nom ro'yxatning qaysi tumanidan olingani
  Izoh                — 2-variant, tuman topilmagani va h.k.

Solishtirish qanday ishlaydi:
  1. Tuman nomlari moslanadi: "Angor tumani" = "Angor" = "Ангор тумани",
     "Termiz shahri" va "Termiz tumani" alohida hisoblanadi.
  2. Mahalla nomi "kalit"ga aylantiriladi: qo'shtirnoq, MFY, tutuq belgilari olib tashlanadi,
     kirill -> lotin, X=H, Q=K, O`=O, G`=G deb qaraladi
     ("XALQOBOD" = "HALKOBOD", "QO`SHARCHA" = "KOSHARCHA").
  3. Faqat o'sha tumandagi mahallalar orasidan eng o'xshashi tanlanadi.
     Raqamlar farq qilsa ("QURAMA-1" va "QURAMA-2"), foiz keskin pasaytiriladi.
  4. Tuman ro'yxatda topilmasa — butun ro'yxatdan qidiriladi va Izohda yoziladi.

"MFY moslik" varag'i ham yaratiladi: har bir noyob nom -> topilgan nom, foiz, soni.
Shu varaqni foiz bo'yicha saralab, sariq/qizillarni tez tekshirib chiqasiz.

Ishga tushirish:
  py mfy_moslash.py
"""

import os
import re
import sys
from copy import copy
from difflib import SequenceMatcher

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.cell_range import CellRange

# ================== ASOSIY SOZLAMALAR ==================
MAIN_FILE = "Yatt o'tmagan 02.11.2026 18-00_mahalla.xlsx"  # topilmasa, ro'yxatdan tanlatadi
REF_FILE = "МФЙ лар руйхати.xlsx"

MAIN_TUMAN_HEADERS = ["tuman", "tuman/shahar", "tuman nomi", "туман", "туман номи"]
MAIN_MFY_HEADERS = ["mahalla nomi", "mahalla", "mfy nomi", "мфй номи"]   # birinchi topilgani ishlatiladi

REF_TUMAN_HEADERS = ["туман номи", "tuman nomi", "туман", "tuman"]
REF_MFY_HEADERS = ["мфй номи", "mfy nomi", "mahalla nomi", "mahalla"]
REF_STIR_HEADERS = ["стир", "stir", "инн", "inn"]

GREEN_FROM = 90      # shundan yuqori — ishonchli (yashil)
YELLOW_FROM = 75     # shu oraliq — sariq; pastda — qizil
MIN_FILL = 90        # bundan past bo'lsa nom QO'YILMAYDI (faqat Izohda taklif)
# =======================================================

NEW_COLS = ["MFY nomi (ro'yxat)", "Moslik %", "MFY STIR", "Ro'yxat tumani", "Izoh"]
FILL = {
    "g": PatternFill("solid", fgColor="C6EFCE"),
    "y": PatternFill("solid", fgColor="FFEB9C"),
    "r": PatternFill("solid", fgColor="FFC7CE"),
}

CYR = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo", "ж": "j", "з": "z",
    "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r",
    "с": "s", "т": "t", "у": "u", "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh",
    "ъ": "", "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "қ": "q", "ғ": "g",
    "ҳ": "h",
}
APOS = "'`‘’ʻʼ´ʹ′"
QUOTES = "\"«»“”„‟″"


def translit(s):
    return "".join(CYR.get(ch, ch) for ch in s.lower())


def base_clean(s):
    s = translit(str(s or "")).replace(" ", " ")
    s = re.sub("[" + re.escape(APOS) + "]", "", s)          # o'z -> oz, g'alaba -> galaba
    s = re.sub("[" + re.escape(QUOTES) + "]", " ", s)
    return s


MFY_NOISE = re.compile(
    r"\b(mahalla\s+fuqarolar(i)?\s+yigini|fuqarolar\s+yigini|m\s*\.?\s*f\s*\.?\s*y\s*\.?(\s*i)?\b"
    r"|mahallasi|mahalla|maxallasi|maxalla|hududida|hududi)\b")


def phon(s):
    """Talaffuz bo'yicha bir xil bo'lgan harflarni birlashtiradi."""
    s = s.replace("h", "x").replace("q", "k").replace("yu", "u").replace("yo", "o")
    s = re.sub(r"(.)\1+", r"\1", s)      # takror harflar: "ALLA" -> "ALA"
    return s


def mfy_key(s):
    s = base_clean(s)
    s = MFY_NOISE.sub(" ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return phon(s)


TUMAN_NOISE = re.compile(r"\b(tumani|tuman|t\.|t)\b")
SHAHAR_NOISE = re.compile(r"\b(shahri|shahar|sh\.|sh)\b")


def tuman_key(s):
    s = base_clean(s)
    city = bool(SHAHAR_NOISE.search(s))
    s = SHAHAR_NOISE.sub(" ", s)
    s = TUMAN_NOISE.sub(" ", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    return phon(s), city


def similarity(a, b):
    """0..100. a, b — mfy_key natijalari."""
    if not a or not b:
        return 0
    if a == b:
        return 100
    a_ns, b_ns = a.replace(" ", ""), b.replace(" ", "")
    if a_ns == b_ns:
        return 99
    r1 = SequenceMatcher(None, a_ns, b_ns).ratio()
    r2 = SequenceMatcher(None, " ".join(sorted(a.split())), " ".join(sorted(b.split()))).ratio()
    score = max(r1, r2) * 100
    # bitta nom ikkinchisining boshida to'liq bo'lsa (MUSTAQILLIK / MUSTAQILLIK 10 YILLIGI)
    if (a_ns.startswith(b_ns) or b_ns.startswith(a_ns)) and min(len(a_ns), len(b_ns)) >= 5:
        score = max(score, 80)
    # raqamlar farq qilsa — boshqa mahalla
    # birinchi harf boshqa bo'lsa — odatda boshqa mahalla (DO`STLIK / YOSHLIK)
    if a_ns[0] != b_ns[0]:
        score *= 0.85
    da, db = re.findall(r"\d+", a), re.findall(r"\d+", b)
    if da != db and (da or db):
        score *= 0.75
    return round(score)


# ---------------- Ro'yxat (reference) ----------------
def find_cols(ws, groups, max_rows=15):
    """groups: {nom: [variantlar]} -> (sarlavha qatori, {nom: ustun})"""
    for r in range(1, min(ws.max_row, max_rows) + 1):
        row = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None:
                row.setdefault(str(v).strip().lower(), c)
        found = {}
        for name, variants in groups.items():
            for v in variants:
                if v in row:
                    found[name] = row[v]
                    break
        if "mfy" in found and "tuman" in found:
            return r, found
    return None, None


def load_reference(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ref = []
    for ws in wb.worksheets:
        hdr, cols = find_cols(ws, {"tuman": REF_TUMAN_HEADERS, "mfy": REF_MFY_HEADERS,
                                   "stir": REF_STIR_HEADERS})
        if not hdr:
            continue
        for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
            name = row[cols["mfy"] - 1]
            if not name:
                continue
            tuman = row[cols["tuman"] - 1]
            stir = row[cols["stir"] - 1] if "stir" in cols else None
            ref.append({"name": str(name).strip(), "tuman": str(tuman or "").strip(),
                        "stir": stir, "key": mfy_key(name), "tkey": tuman_key(tuman)})
    if not ref:
        sys.exit(f"'{path}' faylida 'Туман номи' va 'МФЙ НОМИ' ustunlari topilmadi.")
    return ref


class Matcher:
    def __init__(self, ref):
        self.ref = ref
        self.by_tuman = {}
        for r in ref:
            self.by_tuman.setdefault(r["tkey"], []).append(r)
        self.cache = {}

    def tuman_candidates(self, tuman):
        k, city = tuman_key(tuman)
        if not k:
            return None, "tuman bo'sh"
        if (k, city) in self.by_tuman:
            return self.by_tuman[(k, city)], ""
        same_base = [v for (b, c), v in self.by_tuman.items() if b == k]
        if len(same_base) == 1:
            return same_base[0], ""
        if same_base:      # ham shahar, ham tuman bor — ikkalasidan
            return [x for v in same_base for x in v], ""
        best, bs = None, 0
        for (b, c) in self.by_tuman:
            s = SequenceMatcher(None, k, b).ratio()
            if s > bs:
                best, bs = (b, c), s
        if bs >= 0.8:
            return self.by_tuman[best], ""
        return None, f"'{tuman}' tumani ro'yxatda yo'q — butun ro'yxatdan qidirildi"

    def match(self, tuman, mfy):
        ck = (str(tuman), str(mfy))
        if ck in self.cache:
            return self.cache[ck]
        key = mfy_key(mfy)
        cands, note = self.tuman_candidates(tuman)
        if cands is None:
            cands = self.ref
        scored = sorted(((similarity(key, c["key"]), c) for c in cands),
                        key=lambda x: -x[0])
        res = {"best": None, "score": 0, "note": note, "fill": False}
        if scored:
            s1, c1 = scored[0]
            res.update(best=c1, score=s1, fill=s1 >= MIN_FILL)
            if note:
                res["note"] = f"{note} (topilgani: {c1['tuman']})"
            elif s1 < MIN_FILL:
                # o'z tumanida yo'q — butun ro'yxatda bormi?
                g = max(((similarity(key, c["key"]), c) for c in self.ref), key=lambda x: x[0])
                if g[0] >= MIN_FILL and g[1]["tkey"] != c1["tkey"]:
                    res["note"] = (f"o'z tumanida yo'q; ro'yxatda {g[1]['tuman']} tumanida bor: "
                                   f"{g[1]['name']} ({g[0]}%)")
                else:
                    res["note"] = "ro'yxatda topilmadi"
            if len(scored) > 1:
                s2, c2 = scored[1]
                if s1 < 100 and s2 >= YELLOW_FROM and s1 - s2 <= 10:
                    extra = f"2-variant: {c2['name']} ({s2}%)"
                    res["note"] = (res["note"] + "; " if res["note"] else "") + extra
        self.cache[ck] = res
        return res


# ---------------- Asosiy fayl ----------------
def pick(path, exclude_ref):
    if path and os.path.exists(path):
        return path
    files = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xlsm"))
             and not f.startswith("~$") and f != exclude_ref and "_mfy" not in f]
    if not files:
        sys.exit("Papkada asosiy Excel fayl topilmadi.")
    if len(files) == 1:
        return files[0]
    print("Qaysi faylga MFY nomlarini qo'yamiz?")
    for n, f in enumerate(files, 1):
        print(f"  {n}. {f}")
    while True:
        ch = input("Raqamni kiriting: ").strip()
        if ch.isdigit() and 1 <= int(ch) <= len(files):
            return files[int(ch) - 1]


def insert_cols_after(ws, src_col, n, hdr_row, names):
    new_col = src_col + 1
    merged = [CellRange(str(m)) for m in ws.merged_cells.ranges]
    for m in merged:
        ws.unmerge_cells(str(m))
    widths = {column_index_from_string(l): d.width
              for l, d in list(ws.column_dimensions.items()) if d.width}
    ws.insert_cols(new_col, n)
    for idx in sorted(widths, reverse=True):
        if idx >= new_col:
            ws.column_dimensions[get_column_letter(idx + n)].width = widths[idx]
    for i, w in enumerate([28, 10, 13, 16, 55]):
        ws.column_dimensions[get_column_letter(new_col + i)].width = w
    for m in merged:
        if m.min_col >= new_col:
            m.shift(col_shift=n)
        elif m.max_col >= new_col:
            m.expand(right=n)
        ws.merge_cells(m.coord)
    if ws.auto_filter and ws.auto_filter.ref:
        cr = CellRange(ws.auto_filter.ref)
        if cr.max_col >= src_col:
            cr.expand(right=n)
        ws.auto_filter.ref = cr.coord
    for r in range(1, ws.max_row + 1):
        s = ws.cell(r, src_col)
        if not s.has_style:
            continue
        for i in range(n):
            d = ws.cell(r, new_col + i)
            d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
            d.alignment, d.number_format = copy(s.alignment), s.number_format
    for i, name in enumerate(names):
        ws.cell(hdr_row, new_col + i).value = name
    return new_col


def process_sheet(ws, matcher, summary):
    hdr, cols = find_cols(ws, {"tuman": MAIN_TUMAN_HEADERS, "mfy": MAIN_MFY_HEADERS})
    if not hdr:
        return 0
    # agar oldin qo'shilgan bo'lsa — qayta ishlatamiz
    after = ws.cell(hdr, cols["mfy"] + 1).value
    if after is not None and str(after).strip() == NEW_COLS[0]:
        c0 = cols["mfy"] + 1
        print(f"  [{ws.title}] ustunlar bor edi — yangilanadi")
    else:
        c0 = insert_cols_after(ws, cols["mfy"], len(NEW_COLS), hdr, NEW_COLS)
        print(f"  [{ws.title}] '{ws.cell(hdr, cols['mfy']).value}' ustunidan keyin "
              f"{get_column_letter(c0)}–{get_column_letter(c0 + len(NEW_COLS) - 1)} ustunlar qo'shildi")

    cnt = 0
    for r in range(hdr + 1, ws.max_row + 1):
        mfy = ws.cell(r, cols["mfy"]).value
        tuman = ws.cell(r, cols["tuman"]).value
        cells = [ws.cell(r, c0 + i) for i in range(len(NEW_COLS))]
        if mfy is None or not str(mfy).strip() or str(mfy).strip() == "0":
            for c in cells:
                c.value = None
            continue
        m = matcher.match(tuman, mfy)
        s, best, note = m["score"], m["best"], m["note"]
        if best and m["fill"]:
            cells[0].value, cells[2].value, cells[3].value = best["name"], best["stir"], best["tuman"]
        else:
            cells[0].value = cells[2].value = cells[3].value = None
            if best:
                note = (f"eng yaqini: {best['name']} ({s}%)" + ("; " + note if note else ""))
        cells[1].value = s
        cells[4].value = note or None
        f = FILL["g"] if s >= GREEN_FROM else FILL["y"] if s >= YELLOW_FROM else FILL["r"]
        cells[0].fill = cells[1].fill = f
        cnt += 1
        k = (str(tuman).strip(), str(mfy).strip())
        summary.setdefault(k, {"best": best if m["fill"] else None, "score": s, "note": note, "n": 0})
        summary[k]["n"] += 1
    return cnt


def write_summary(wb, summary):
    name = "MFY moslik"
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.append(["Tuman", "Asl nom", "Ro'yxatdagi nom", "Moslik %", "STIR", "Ro'yxat tumani", "Soni", "Izoh"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for (t, n), v in sorted(summary.items(), key=lambda x: (x[1]["score"], x[0])):
        b = v["best"]
        ws.append([t, n, b["name"] if b else "", v["score"], b["stir"] if b else "",
                   b["tuman"] if b else "", v["n"], v["note"]])
        s = v["score"]
        ws.cell(ws.max_row, 4).fill = FILL["g"] if s >= GREEN_FROM else FILL["y"] if s >= YELLOW_FROM else FILL["r"]
    for col, w in zip("ABCDEFGH", [18, 30, 30, 10, 13, 16, 7, 60]):
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = f"A1:H{ws.max_row}"
    ws.freeze_panes = "A2"


def main():
    if not os.path.exists(REF_FILE):
        sys.exit(f"'{REF_FILE}' topilmadi. Uni skript bilan bir papkaga qo'ying.")
    main_file = pick(sys.argv[1] if len(sys.argv) > 1 else MAIN_FILE, REF_FILE)
    out = os.path.splitext(main_file)[0] + "_mfy.xlsx"

    print(f"Ro'yxat: {REF_FILE}")
    ref = load_reference(REF_FILE)
    print(f"  {len(ref)} ta MFY, {len({r['tkey'] for r in ref})} ta tuman/shahar")
    print(f"Asosiy fayl: {main_file}")
    wb = load_workbook(main_file)
    matcher, summary, total = Matcher(ref), {}, 0
    for ws in wb.worksheets:
        if ws.title in ("MFY moslik", "Tekshiruv"):
            continue
        total += process_sheet(ws, matcher, summary)
    if not total:
        sys.exit("Asosiy faylda 'Tuman' va 'Mahalla' ustunlari topilmadi.")
    write_summary(wb, summary)
    try:
        wb.save(out)
    except PermissionError:
        sys.exit(f"'{out}' ochiq turibdi — Excel'da yopib, qayta ishga tushiring.")

    rows = [v for v in summary.values() for _ in range(v["n"])]
    g = sum(1 for v in rows if v["score"] >= GREEN_FROM)
    y = sum(1 for v in rows if YELLOW_FROM <= v["score"] < GREEN_FROM)
    print(f"\nSaqlandi: {out}")
    print(f"  Jami {total} qator: yashil (>= {GREEN_FROM}%) {g}, sariq {y}, qizil {total - g - y}")
    print("  Sariq/qizillarni 'MFY moslik' varag'ida tekshiring.")


if __name__ == "__main__":
    main()
