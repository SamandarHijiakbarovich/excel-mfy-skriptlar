# -*- coding: utf-8 -*-
"""
YURIDIK o'tmaganlar fayliga "МФЙ лар руйхати.xlsx" dan to'g'ri MFY nomini qo'yadi
faqat SOATO (tuman) + mahalla NOMI bo'yicha (MFY INN ustuni hisobga olinmaydi)
(masalan, Sherobodning mahallasi Denovga tushib qolmasligi uchun).

"MFY nomi" ustunidan keyin 4 ta ustun qo'shiladi (boshqa ustunlar joyida qoladi):
  MFY nomi (ro'yxat)  — "KO`KLAM" MFY ko'rinishida (kechagi stil)
  MFY STIR (ro'yxat)  — ro'yxatdagi to'g'ri СТИР
  Tekshiruv           — OK / TEKSHIRING / XATO
  Izoh                — nima uchun (STIR boshqa tumanga tegishli, nom farq qiladi va h.k.)

Har bir qator qanday tekshiriladi:
  1. Qatordagi "MFY INN" ro'yxatdagi СТИР dan qidiriladi.
     a) Topildi va ro'yxatdagi МХОБТ == qatordagi SOATO  -> to'g'ri (nom ham solishtiriladi)
     b) Topildi, lekin МХОБТ boshqa (boshqa tuman mahallasi!) -> o'z tumanidan nom bo'yicha qidiriladi
  2. MFY INN ro'yxatda yo'q / bo'sh -> o'z tumani (SOATO) ichidan "MFY nomi" bo'yicha qidiriladi.
  3. "MFY nomi" ham yordam bermasa -> "Manzil"dan mahalla nomi ajratib olinadi
     ("... Ko'klam mahallasi, ..." -> KO`KLAM) va o'z tumanidan qidiriladi.
  Nom bo'yicha qidiruvda faqat 90% va undan yuqori moslik qabul qilinadi.
  Hech qachon boshqa SOATO dagi mahalla qo'yilmaydi.

Excel jadvali (Format as Table), formulalar, filtr, shartli formatlash surilgan ustunlarga moslab yangilanadi.
Natija yangi faylga yoziladi: <nom>_mfy.xlsx (asl fayl o'zgarmaydi).
"MFY tekshiruv" varag'ida faqat muammoli qatorlar ro'yxati bo'ladi.

Ishga tushirish:
  py -m pip install openpyxl
  py yuridik_mfy.py
"""

import os
import re
import sys
from copy import copy
from difflib import SequenceMatcher

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.table import TableColumn

# ================== ASOSIY SOZLAMALAR ==================
MAIN_FILE = "ЮРИДИК утмаган 02.11.2026 18-00.xlsx"   # topilmasa, ro'yxatdan tanlatadi
REF_FILE = "МФЙ лар руйхати.xlsx"

H_SOATO = ["soato", "соато", "мхобт", "mxobt"]
H_MFY_INN = ["mfy inn", "mfy stir", "мфй инн", "мфй стир"]
H_MFY_NAME = ["mfy nomi", "мфй номи", "mahalla", "mahalla nomi"]
H_TUMAN = ["tuman/shahar", "tuman", "tuman nomi", "туман номи"]
H_MANZIL = ["manzil", "манзил"]

R_STIR = ["стир", "stir", "инн", "inn"]
R_TUMAN = ["туман номи", "tuman nomi", "туман", "tuman"]
R_NAME = ["мфй номи", "mfy nomi", "mahalla nomi"]
R_SOATO = ["мхобт", "mxobt", "soato", "соато"]

MIN_NAME = 90          # nom bo'yicha qidiruvda qabul qilinadigan eng kichik foiz
NAME_WARN = 80         # STIR to'g'ri, lekin nom shundan kam o'xshasa — "TEKSHIRING"
APOSTROPHE = "`"       # o`, g` belgisi
# =======================================================

NEW_COLS = ["MFY nomi (ro'yxat)", "MFY STIR (ro'yxat)", "Tekshiruv", "Izoh"]
FILL = {"OK": PatternFill("solid", fgColor="C6EFCE"),
        "TEKSHIRING": PatternFill("solid", fgColor="FFEB9C"),
        "XATO": PatternFill("solid", fgColor="FFC7CE")}

# ---------------- Nomni kechagi stilga keltirish: "NOM" MFY ----------------
APOS_CHARS = "'`‘’ʻʼ´ʹ′"
QUOTE_CHARS = "\"«»“”„‟″"
PH = "\u0001"
_NOISE = [
    r"mahalla\s+fuqarolar(?:i)?\s+yig\W?ini", r"fuqarolar\s+yig\W?ini",
    r"махалла\s+фуқаролар(?:и)?\s+йиғини", r"маҳалла\s+фуқаролар(?:и)?\s+йиғини",
    r"фуқаролар\s+йиғини", r"m\s*\.?\s*f\s*\.?\s*y\s*\.?(?:\s*i)?", r"м\s*\.?\s*ф\s*\.?\s*й\s*\.?(?:\s*и)?",
    r"mahallasi", r"mahalla", r"maxallasi", r"maxalla", r"маҳалласи", r"махалласи", r"маҳалла", r"махалла",
]
NOISE_RE = re.compile(r"(?<![\w" + PH + r"])(?:" + "|".join(_NOISE) + r")(?![\w" + PH + r"])", re.I)


def style_name(value):
    if value is None:
        return ""
    s = str(value).replace(" ", " ").strip()
    if not s or s in ("0", "-"):
        return ""
    s = re.sub(r"(?<=\w)[" + re.escape(APOS_CHARS) + r"](?=\w)", PH, s)
    s = re.sub(r"(?<=[oOgG])[" + re.escape(APOS_CHARS) + r"](?=\s|$)", PH, s)
    s = re.sub("[" + re.escape(QUOTE_CHARS + APOS_CHARS) + "]", " ", s)
    s = NOISE_RE.sub(" ", s)
    s = re.sub(r"[,;:()\[\]]+", " ", s)
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s+", " ", s).strip(" .-").replace(PH, APOSTROPHE)
    return f'"{s.upper()}" MFY' if s else ""


# ---------------- Solishtirish kaliti ----------------
CYR = dict(zip("абвгдежзийклмнопрстуфхцчшщъыьэюяўқғҳё",
               ["a", "b", "v", "g", "d", "e", "j", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s",
                "t", "u", "f", "x", "ts", "ch", "sh", "sh", "", "i", "", "e", "yu", "ya", "o", "q", "g",
                "h", "yo"]))


def mfy_key(s):
    s = "".join(CYR.get(c, c) for c in str(s or "").lower())
    s = re.sub("[" + re.escape(APOS_CHARS) + "]", "", s)
    s = re.sub("[" + re.escape(QUOTE_CHARS) + "]", " ", s)
    s = re.sub(r"\b(mahalla\s+fuqarolar(i)?\s+yigini|fuqarolar\s+yigini|m\s*\.?\s*f\s*\.?\s*y\s*\.?(\s*i)?\b"
               r"|mahallasi|mahalla|maxallasi|maxalla)\b", " ", s)
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace("h", "x").replace("q", "k").replace("yu", "u").replace("yo", "o")
    return re.sub(r"(.)\1+", r"\1", s)


def coarse(k):
    """Unlilar farqini yo'qotadigan qo'pol kalit: o/a/u -> a, i/y/e -> i
    (FAROVON=FARAVON, XOJAIPOK=XO`JAYPOK, URGULSOY=O`RG`ULSOY, TALLOSHQON=TALLASHQON)."""
    k = k.replace(" ", "").replace("-", "")
    k = re.sub(r"[aou]", "a", k)
    k = re.sub(r"[iye]", "i", k)
    k = k.replace("dj", "j").replace("ts", "s")
    return re.sub(r"(.)\1+", r"\1", k)


def similarity(a, b):
    s1 = _sim(a, b)
    if s1 >= 99 or not a or not b:
        return s1
    ca, cb = coarse(a), coarse(b)
    if ca == cb and re.findall(r"\d+", a) == re.findall(r"\d+", b):
        return max(s1, 97)
    return max(s1, round(_sim(ca, cb) * 0.95))


def _sim(a, b):
    if not a or not b:
        return 0
    if a == b:
        return 100
    an, bn = a.replace(" ", ""), b.replace(" ", "")
    if an == bn:
        return 99
    sc = max(SequenceMatcher(None, an, bn).ratio(),
             SequenceMatcher(None, " ".join(sorted(a.split())), " ".join(sorted(b.split()))).ratio()) * 100
    if (an.startswith(bn) or bn.startswith(an)) and min(len(an), len(bn)) >= 5:
        sc = max(sc, 80)
    if an[0] != bn[0]:
        sc *= 0.85
    if re.findall(r"\d+", a) != re.findall(r"\d+", b):
        sc *= 0.75
    return round(sc)


def soato7(v):
    d = re.sub(r"\D", "", str(v or "").split(".")[0])
    return d[:7] if len(d) >= 7 else ""


def stir_norm(v):
    d = re.sub(r"\D", "", str(v or "").split(".")[0])
    return d if len(d) == 9 else ""


def tuman_key(s):
    s = mfy_key(s)
    city = bool(re.search(r"\b(sxaxri|sxaxar|sx)\b", s))
    s = re.sub(r"\b(tumani|tuman|t|sxaxri|sxaxar|sx)\b", " ", s)
    return s.replace(" ", ""), city


MANZIL_RE = re.compile(r"([^,]*?)\s*(?:mahallasi|mahalla|maxallasi|m\.?f\.?y\.?|mfy)\b", re.I)


def names_from_manzil(manzil):
    """Manzildan ehtimoliy mahalla nomlari: '..., Sho'rchi shahri Ko'klam mahallasi, ...' -> ["Ko'klam", ...]"""
    out = []
    for m in MANZIL_RE.finditer(str(manzil or "")):
        words = m.group(1).strip().split()
        for n in (1, 2, 3):
            if len(words) >= n:
                out.append(" ".join(words[-n:]))
    return out


# ---------------- Ro'yxat ----------------
def find_header(ws, groups, required):
    for r in range(1, min(ws.max_row, 15) + 1):
        row = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None:
                row.setdefault(re.sub(r"\s+", " ", str(v)).strip().lower(), c)
        found = {}
        for k, variants in groups.items():
            for v in variants:
                if v in row:
                    found[k] = row[v]
                    break
        if all(k in found for k in required):
            return r, found
    return None, None


class Reference:
    def __init__(self, path):
        wb = load_workbook(path, data_only=True, read_only=True)
        self.by_stir, self.by_soato, self.soato_by_tuman = {}, {}, {}
        n = 0
        for ws in wb.worksheets:
            hdr, c = find_header(ws, {"stir": R_STIR, "tuman": R_TUMAN, "name": R_NAME, "soato": R_SOATO},
                                 ["stir", "name", "soato"])
            if not hdr:
                continue
            for row in ws.iter_rows(min_row=hdr + 1, values_only=True):
                name = row[c["name"] - 1]
                soato = soato7(row[c["soato"] - 1])
                if not name or not soato:
                    continue
                rec = {"stir": stir_norm(row[c["stir"] - 1]), "name": str(name).strip(),
                       "soato": soato, "tuman": str(row[c["tuman"] - 1] or "").strip() if "tuman" in c else "",
                       "key": mfy_key(name)}
                rec["styled"] = style_name(rec["name"])
                if rec["stir"]:
                    self.by_stir[rec["stir"]] = rec
                self.by_soato.setdefault(soato, []).append(rec)
                if rec["tuman"]:
                    self.soato_by_tuman.setdefault(tuman_key(rec["tuman"]), soato)
                n += 1
        if not n:
            sys.exit(f"'{path}' da СТИР / МФЙ НОМИ / МХОБТ ustunlari topilmadi.")
        self.tuman_names = {s: v[0]["tuman"] for s, v in self.by_soato.items()}
        print(f"  ro'yxat: {n} ta MFY, {len(self.by_soato)} ta tuman/shahar (SOATO)")
        self.cache = {}
        self.last_second = (0, None)

    def soato_from_tuman(self, tuman):
        k = tuman_key(tuman)
        if k in self.soato_by_tuman:
            return self.soato_by_tuman[k]
        same = [s for (b, c), s in self.soato_by_tuman.items() if b == k[0]]
        return same[0] if len(same) == 1 else ""

    def best_in_soato(self, soato, names):
        """names — bir nechta nomzod nom; o'sha SOATO ichidan eng yaxshisi."""
        best, bs = None, 0
        for nm in names:
            k = mfy_key(nm)
            if not k:
                continue
            ck = (soato, k)
            if ck not in self.cache:
                cands = self.by_soato.get(soato, [])
                sc = sorted(((similarity(k, c["key"]), c) for c in cands), key=lambda x: -x[0])
                top = sc[0] if sc else (0, None)
                # ikkinchi, BOSHQA nomli eng yaqin variant
                second = next(((x, c) for x, c in sc[1:] if c["key"] != top[1]["key"]), (0, None)) if sc else (0, None)
                self.cache[ck] = (top[0], top[1], second[0], second[1])
            s, c, s2, c2 = self.cache[ck]
            if s > bs:
                best, bs = c, s
                self.last_second = (s2, c2)
        return best, bs

    def tname(self, soato):
        return self.tuman_names.get(soato, soato)


def check_row(ref, soato, mfy_inn, mfy_name, manzil):
    """SOATO + mahalla NOMI bo'yicha (MFY INN hisobga olinmaydi). -> (rec | None, holat, izoh)"""
    if soato not in ref.by_soato:
        return None, "XATO", f"SOATO {soato or '(bo`sh)'} ro'yxatda yo'q"

    ref.last_second = (0, None)
    best, s = ref.best_in_soato(soato, [mfy_name] if mfy_name else [])
    if best and s >= MIN_NAME:
        s2, c2 = ref.last_second
        if c2 and s2 >= MIN_NAME and s - s2 <= 3:
            return best, "TEKSHIRING", f"2 ta o'xshash: {best['styled']} ({s}%) va {c2['styled']} ({s2}%)"
        return best, "OK", "" if s == 100 else f"MFY nomi bo'yicha {s}%"

    best2, s2 = ref.best_in_soato(soato, names_from_manzil(manzil))
    if best2 and s2 >= MIN_NAME:
        why = "MFY nomi bo'sh" if not mfy_name else f"MFY nomi bo'yicha topilmadi ({s}%)"
        return best2, "TEKSHIRING", f"{why}; Manzil bo'yicha topildi ({s2}%)"

    cand, sc = (best, s) if (best and s >= s2) else (best2, s2)
    izoh = "MFY nomi bo'sh, Manzilda ham yo'q" if not mfy_name else f"'{mfy_name}' {ref.tname(soato)} ro'yxatida topilmadi"
    if cand:
        izoh += f"; eng yaqini: {cand['styled']} ({sc}%) — qabul qilinmadi"
    # boshqa tumanlarda bormi? (qo'yilmaydi, faqat ma'lumot)
    if mfy_name:
        other = []
        for so in ref.by_soato:
            if so == soato:
                continue
            ob, os_ = ref.best_in_soato(so, [mfy_name])
            if ob and os_ >= MIN_NAME:
                other.append(f"{ref.tname(so)} ({so})")
        if other:
            izoh += "; ro'yxatda boshqa tumanda bor: " + ", ".join(other[:3]) + " — SOATO ni tekshiring"
    return None, "XATO", izoh


# ---------------- Ustun qo'shish (jadval/formula/filtrni buzmasdan) ----------------
def shift_formula(f, new_col, n, sheet_title):
    if not isinstance(f, str) or not f.startswith("="):
        return f
    tok = Tokenizer(f)
    out = []
    ref_re = re.compile(r"(\$?)([A-Z]{1,3})(?=\$?\d|:|$)")

    for t in tok.items:
        v = t.value
        if t.type == "OPERAND" and t.subtype == "RANGE":
            sheet, rng = (v.rsplit("!", 1) if "!" in v else (None, v))
            same = sheet is None or sheet.strip("'") == sheet_title
            if same and not rng.startswith("[") and "#" not in rng:
                def rep(m):
                    ci = column_index_from_string(m.group(2))
                    return m.group(1) + get_column_letter(ci + n if ci >= new_col else ci)
                rng = re.sub(r"(\$?)([A-Z]{1,3})(?=\$?\d|:\$?[A-Z]|$)", rep, rng)
                v = (sheet + "!" if sheet else "") + rng
        out.append(v)
    return "=" + "".join(out)


def insert_cols(ws, after_col, names):
    n, new_col = len(names), after_col + 1
    title = ws.title

    merged = [CellRange(str(m)) for m in ws.merged_cells.ranges]
    for m in merged:
        ws.unmerge_cells(str(m))
    widths = {column_index_from_string(l): d.width for l, d in list(ws.column_dimensions.items()) if d.width}
    hidden = {column_index_from_string(l) for l, d in list(ws.column_dimensions.items()) if d.hidden}
    cf_old = [(str(cf.sqref), cf.rules) for cf in ws.conditional_formatting]

    ws.insert_cols(new_col, n)

    # formulalar
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                try:
                    c.value = shift_formula(c.value, new_col, n, title)
                except Exception:
                    pass
    # kenglik / yashirin ustunlar
    for idx in sorted(widths, reverse=True):
        if idx >= new_col:
            ws.column_dimensions[get_column_letter(idx + n)].width = widths[idx]
    for idx in hidden:
        if idx >= new_col:
            ws.column_dimensions[get_column_letter(idx)].hidden = False
            ws.column_dimensions[get_column_letter(idx + n)].hidden = True
    for i, w in enumerate([26, 14, 13, 60]):
        ws.column_dimensions[get_column_letter(new_col + i)].width = w

    def shift_range(coord):
        cr = CellRange(coord)
        if cr.min_col >= new_col:
            cr.shift(col_shift=n)
        elif cr.max_col >= new_col - 1:
            cr.expand(right=n)
        return cr.coord

    def span_shift(coord):
        cr = CellRange(coord)
        if cr.min_col >= new_col:
            cr.shift(col_shift=n)
        elif cr.max_col >= new_col:
            cr.expand(right=n)
        return cr.coord

    for m in merged:
        ws.merge_cells(span_shift(m.coord))
    if ws.auto_filter and ws.auto_filter.ref:
        ws.auto_filter.ref = shift_range(ws.auto_filter.ref)
    # shartli formatlash
    ws.conditional_formatting = ConditionalFormattingList()
    for sq, rules in cf_old:
        new_sq = " ".join(span_shift(p) for p in sq.split())
        for r in rules:
            ws.conditional_formatting.add(new_sq, r)
    # Excel jadvali (Format as Table)
    for t in ws.tables.values():
        cr = CellRange(t.ref)
        if cr.min_col <= after_col <= cr.max_col:
            pos = after_col - cr.min_col + 1
            cr.expand(right=n)
            t.ref = cr.coord
            if t.autoFilter is not None:
                t.autoFilter.ref = cr.coord
            cols = list(t.tableColumns)
            for i, nm in enumerate(names):
                cols.insert(pos + i, TableColumn(id=0, name=nm))
            for i, col in enumerate(cols, 1):
                col.id = i
            t.tableColumns = cols
        elif cr.min_col >= new_col:
            cr.shift(col_shift=n)
            t.ref = cr.coord
            if t.autoFilter is not None:
                t.autoFilter.ref = cr.coord

    # stil "MFY nomi" ustunidan
    for r in range(1, ws.max_row + 1):
        s = ws.cell(r, after_col)
        if s.has_style:
            for i in range(n):
                d = ws.cell(r, new_col + i)
                d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
                d.alignment = copy(s.alignment)
    return new_col


# ---------------- Asosiy ----------------
def pick(path):
    if path and os.path.exists(path):
        return path
    files = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xlsm"))
             and not f.startswith("~$") and f != REF_FILE and "_mfy" not in f]
    if not files:
        sys.exit("Papkada Excel fayl topilmadi.")
    if len(files) == 1:
        return files[0]
    print("Qaysi faylni ishlaymiz?")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    while True:
        ch = input("Raqamni kiriting: ").strip()
        if ch.isdigit() and 1 <= int(ch) <= len(files):
            return files[int(ch) - 1]


def main():
    if not os.path.exists(REF_FILE):
        sys.exit(f"'{REF_FILE}' topilmadi — skript bilan bir papkaga qo'ying.")
    main_file = pick(sys.argv[1] if len(sys.argv) > 1 else MAIN_FILE)
    out = os.path.splitext(main_file)[0] + "_mfy.xlsx"

    print(f"Ro'yxat: {REF_FILE}")
    ref = Reference(REF_FILE)
    print(f"Ochilmoqda: {main_file} (katta fayl bo'lsa biroz kutasiz)")
    wb = load_workbook(main_file)
    problems, stats = [], {"OK": 0, "TEKSHIRING": 0, "XATO": 0}

    for ws in wb.worksheets:
        if ws.title == "MFY tekshiruv":
            continue
        hdr, c = find_header(ws, {"soato": H_SOATO, "inn": H_MFY_INN, "name": H_MFY_NAME,
                                  "tuman": H_TUMAN, "manzil": H_MANZIL}, ["name"])
        if not hdr or not ({"soato", "tuman"} & set(c)):
            continue
        nxt = ws.cell(hdr, c["name"] + 1).value
        if nxt is not None and str(nxt).strip() == NEW_COLS[0]:
            c0 = c["name"] + 1
        else:
            c0 = insert_cols(ws, c["name"], NEW_COLS)
            c = {k: (v + len(NEW_COLS) if v > c["name"] else v) for k, v in c.items()}
        for i, nm in enumerate(NEW_COLS):
            ws.cell(hdr, c0 + i).value = nm
            ws.cell(hdr, c0 + i).font = copy(ws.cell(hdr, c["name"]).font)
        print(f"  [{ws.title}] ustunlar: {get_column_letter(c0)}–{get_column_letter(c0 + 3)}")

        g = lambda r, k: ws.cell(r, c[k]).value if k in c else None
        for r in range(hdr + 1, ws.max_row + 1):
            name, inn, soato_v, tuman = g(r, "name"), g(r, "inn"), g(r, "soato"), g(r, "tuman")
            if all(v in (None, "") for v in (name, inn, soato_v, g(r, "manzil"))):
                continue
            soato = soato7(soato_v) or ref.soato_from_tuman(tuman)
            rec, holat, izoh = check_row(ref, soato, stir_norm(inn), str(name or "").strip(), g(r, "manzil"))
            cells = [ws.cell(r, c0 + i) for i in range(4)]
            cells[0].value = rec["styled"] if rec else None
            cells[1].value = int(rec["stir"]) if rec and rec["stir"] else None
            cells[2].value = holat
            cells[3].value = izoh or None
            cells[0].fill = cells[2].fill = FILL[holat]
            stats[holat] += 1
            if holat != "OK":
                problems.append([ws.title, r, tuman, soato, inn, name, rec["styled"] if rec else "",
                                 rec["stir"] if rec else "", holat, izoh])

    if not sum(stats.values()):
        sys.exit("Faylda 'MFY nomi' va 'SOATO' ustunlari topilmadi.")

    if "MFY tekshiruv" in wb.sheetnames:
        del wb["MFY tekshiruv"]
    rs = wb.create_sheet("MFY tekshiruv")
    rs.append(["Varaq", "Qator", "Tuman", "SOATO", "MFY INN (asl)", "MFY nomi (asl)",
               "MFY nomi (ro'yxat)", "MFY STIR (ro'yxat)", "Tekshiruv", "Izoh"])
    for x in rs[1]:
        x.font = Font(bold=True)
    for p in problems:
        rs.append(p)
        rs.cell(rs.max_row, 9).fill = FILL[p[8]]
    for col, w in zip("ABCDEFGHIJ", [12, 7, 14, 10, 13, 24, 26, 14, 13, 90]):
        rs.column_dimensions[col].width = w
    rs.auto_filter.ref = f"A1:J{max(rs.max_row, 2)}"
    rs.freeze_panes = "A2"

    try:
        wb.save(out)
    except PermissionError:
        sys.exit(f"'{out}' ochiq turibdi — yopib, qayta ishga tushiring.")
    print(f"\nSaqlandi: {out}")
    print(f"  OK: {stats['OK']}   TEKSHIRING: {stats['TEKSHIRING']}   XATO: {stats['XATO']}")
    print("  Muammoli qatorlar 'MFY tekshiruv' varag'ida.")


if __name__ == "__main__":
    main()
