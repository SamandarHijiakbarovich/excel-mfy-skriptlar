# -*- coding: utf-8 -*-
"""
"солик" varag'idagi F.I.Sh.larni (kirill) lotinga o'giradi va "adliya" varag'idagi
F.I.Sh.lar bilan solishtirib, mos kelganini yoniga olib o'tadi.

"солик" varag'iga F.I.Sh. ustunidan keyin qo'shiladigan ustunlar:
  ФИШ (lotin)   — kirilldan o'girilgan
  Adliya ИФШ    — adliyada topilgan F.I.Sh. (asl ko'rinishida)
  Moslik %      — faqat 100% (aniq) mos kelganlar olib o'tiladi; taxminiylari Izohda
  Izoh          — "faqat familiya-ism mos", "adliyada 2 ta" va h.k.
"adliya" varag'iga "Soliqda bor" ustuni qo'shiladi (qaysi qatorga mos kelgani).

Solishtirishda yozuv farqlari hisobga olinmaydi:
  O'G'LI = O‘G‘LI = OGLI = UGLI = ЎҒЛИ,  QIZI = KIZI,  X = H,  Q = K,
  ABDULLAYEVNA = ABDULLAEVNA,  DJURABAYEVNA = JO'RABOYEVNA,  "XXX" e'tiborga olinmaydi.

Natija: <nom>_natija.xlsx (asl fayl o'zgarmaydi)

Ishga tushirish:
  py -m pip install openpyxl
  py fish_moslash.py
"""

import os
import re
import sys
from collections import defaultdict
from difflib import SequenceMatcher

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================== ASOSIY SOZLAMALAR ==================
INPUT_FILE = "Книга6.xlsx"
SOLIQ_SHEET = ["солик", "соliq", "soliq", "солиқ"]      # varaq nomi variantlari
ADLIYA_SHEET = ["adliya", "адлия", "adliy"]
NAME_HEADERS = ["фиш", "ифш", "fish", "ifsh", "f.i.sh", "ф.и.ш"]  # sarlavhada shu so'z bo'lsa
APOSTROPHE = "'"          # lotin yozuvida o', g' uchun: "'" yoki "‘"
# Faqat ANIQ mos kelganlar olib o'tiladi (100%): yozuv farqlari (O'G'LI/O‘G‘LI/ЎҒЛИ, Q/K, X/H,
# DJ/J, YE/E, XXX) hisobga olinmaydi, lekin bitta harf ham boshqa bo'lsa — olib o'tilmaydi.
# Taxminiy o'xshashlar faqat "Izoh"da ko'rsatiladi. Taxminiylarni ham olish uchun: MIN_ACCEPT = 95
MIN_ACCEPT = 100
GREEN_FROM = 100
# =======================================================

NEW_COLS = ["ФИШ (lotin)", "Adliya ИФШ", "Moslik %", "Izoh"]
FILL = {"g": PatternFill("solid", fgColor="C6EFCE"),
        "y": PatternFill("solid", fgColor="FFEB9C"),
        "r": PatternFill("solid", fgColor="FFC7CE")}

# ---------------- Kirill -> lotin ----------------
A = APOSTROPHE
MAP = {
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Ж": "J", "З": "Z", "И": "I", "Й": "Y",
    "К": "K", "Л": "L", "М": "M", "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T",
    "У": "U", "Ф": "F", "Х": "X", "Ц": "S", "Ч": "CH", "Ш": "SH", "Щ": "SH", "Ъ": A, "Ы": "I",
    "Ь": "", "Э": "E", "Ю": "YU", "Я": "YA", "Ё": "YO", "Ў": "O" + A, "Қ": "Q", "Ғ": "G" + A,
    "Ҳ": "H", "Е": "E",
}
VOWELS = set("АЕЁИОУЎЭЮЯЫ")


def to_latin(s):
    s = str(s or "").strip().upper()
    out = []
    for i, ch in enumerate(s):
        if ch == "Е":
            prev = s[i - 1] if i else " "
            # so'z boshida yoki unli/Ъ/Ь dan keyin -> YE
            out.append("YE" if (not prev.isalpha() or prev in VOWELS or prev in "ЪЬ") else "E")
        elif ch in MAP:
            out.append(MAP[ch])
        else:
            out.append(ch)
    r = "".join(out)
    r = re.sub(r"[‘’ʻʼ`´]", A, r)
    return re.sub(r"\s+", " ", r).strip()


# ---------------- Solishtirish kaliti ----------------
def tokens(name):
    """Lotin F.I.Sh. -> normallashtirilgan so'zlar ro'yxati."""
    s = to_latin(name)                                   # lotin bo'lsa ham apostrof/katta harf bir xil
    s = re.sub(r"['‘’ʻʼ`´]", "", s)                      # O'G'LI -> OGLI
    s = re.sub(r"[^A-Z\s-]", " ", s).replace("-", " ")
    words = [w for w in s.split() if w not in ("XXX", "XX", "X")]
    res = []
    for w in words:
        if w in ("OGLI", "UGLI", "OGLY", "UGLY", "OGL"):
            w = "OGLI"
        elif w in ("QIZI", "KIZI", "QIZ", "KIZ"):
            w = "KIZI"
        w = w.replace("CH", "Ç").replace("SH", "Ş")      # CH/SH ichidagi H saqlanadi
        w = w.replace("H", "X").replace("Q", "K").replace("DJ", "J").replace("YE", "E")
        w = re.sub(r"(.)\1+", r"\1", w)                  # ABDULLA -> ABDULA
        res.append(w)
    # "O'G'LI"/"QIZI" bir so'z bo'lib yozilgan bo'lsa: "BOTIROGLI"
    return res


def coarse(w):
    w = re.sub(r"[OAU]", "A", w)
    w = re.sub(r"[IYE]", "I", w)
    return re.sub(r"(.)\1+", r"\1", w)


SUFFIX_RE = re.compile(r"(OVIÇ|EVIÇ|OVNA|EVNA|IVIÇ|IVNA|OVA|EVA)$")


def _close(w1, w2, strict):
    """Bitta so'z (familiya/ism) bir xilmi: unlilar farqi hisobga olinmaydi."""
    c1, c2 = coarse(w1), coarse(w2)
    if c1 == c2:
        return 1.0
    r = SequenceMatcher(None, c1, c2).ratio()
    if min(len(c1), len(c2)) < 5 or c1[0] != c2[0]:     # qisqa yoki boshqa harfdan boshlangan: ALI / VALI
        return 0.0 if strict else r * 0.5
    return r


def patronym(t):
    """Otasining ismi o'zagi: BOTIR O'G'LI -> BOTIR, BOTIROVICH -> BOTIR, ABDUMALIKOVNA -> ABDUMALIK"""
    if len(t) < 3:
        return ""
    return SUFFIX_RE.sub("", t[2]) if t[2] not in ("OGLI", "KIZI") else ""


def score(t1, t2):
    """Ikki F.I.Sh. (tokens) o'xshashligi 0..100 va izoh."""
    if not t1 or not t2:
        return 0, ""
    if t1 == t2:
        return 100, ""
    if [coarse(w) for w in t1] == [coarse(w) for w in t2]:
        return 97, ""
    if len(t1) < 2 or len(t2) < 2:
        return 0, ""
    fam = _close(t1[0], t2[0], strict=False)
    ism = _close(t1[1], t2[1], strict=True)
    if fam < 0.85 or ism < 0.85:
        return round(min(fam, ism) * 60), ""
    sc = round((fam + ism) / 2 * 100)
    p1, p2 = patronym(t1), patronym(t2)
    if p1 and p2:
        pr = 1.0 if coarse(p1) == coarse(p2) else SequenceMatcher(None, coarse(p1), coarse(p2)).ratio()
        if pr < 0.85 or coarse(p1)[0] != coarse(p2)[0]:
            return min(sc, 80), "familiya va ism mos, otasining ismi boshqa"
        sc = round(sc * (0.9 + 0.1 * pr))
        if pr < 1.0 or len(t1) != len(t2):
            sc = min(sc, 94)
        return sc, ""
    return min(sc, 88), "otasining ismi bir tomonda yo'q"


# ---------------- Excel ----------------
def find_sheet(wb, names):
    for ws in wb.worksheets:
        t = ws.title.strip().lower()
        if any(n in t for n in names):
            return ws
    return None


def find_name_col(ws):
    for r in range(1, min(ws.max_row, 10) + 1):
        for c in range(1, ws.max_column + 1):
            v = str(ws.cell(r, c).value or "").strip().lower().replace(" ", "")
            if any(h.replace(" ", "") in v for h in NAME_HEADERS):
                return r, c
    return 1, 1


def main():
    inp = sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE
    if not os.path.exists(inp):
        files = [f for f in os.listdir(".") if f.lower().endswith(".xlsx")
                 and not f.startswith("~$") and "_natija" not in f]
        if len(files) == 1:
            inp = files[0]
        else:
            sys.exit(f"'{inp}' topilmadi. Skriptni Excel fayl bilan bir papkaga qo'ying.")
    out = os.path.splitext(inp)[0] + "_natija.xlsx"
    print(f"Ochilmoqda: {inp}")
    wb = load_workbook(inp)

    ws_s = find_sheet(wb, SOLIQ_SHEET) or wb.worksheets[0]
    ws_a = find_sheet(wb, ADLIYA_SHEET) or wb.worksheets[1]
    hs, cs = find_name_col(ws_s)
    ha, ca = find_name_col(ws_a)
    print(f"  soliq:  '{ws_s.title}', ustun {get_column_letter(cs)}, sarlavha {hs}-qator")
    print(f"  adliya: '{ws_a.title}', ustun {get_column_letter(ca)}, sarlavha {ha}-qator")

    # --- adliya indeksi ---
    adl = []
    for r in range(ha + 1, ws_a.max_row + 1):
        v = ws_a.cell(r, ca).value
        if v and str(v).strip():
            t = tokens(v)
            if t:
                adl.append({"row": r, "name": str(v).strip(), "t": t, "key": " ".join(t)})
    by_key = defaultdict(list)
    by_block = defaultdict(list)          # familiyaning boshi bo'yicha guruh — tez qidirish uchun
    for a in adl:
        by_key[a["key"]].append(a)
        by_block[coarse(a["t"][0])[:3]].append(a)
    print(f"  adliyada {len(adl)} ta F.I.Sh.")

    # --- soliq varag'iga ustunlar ---
    ws_s.insert_cols(cs + 1, len(NEW_COLS))
    head = ws_s.cell(hs, cs)
    for i, n in enumerate(NEW_COLS):
        c = ws_s.cell(hs, cs + 1 + i, n)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if head.fill and head.fill.fgColor and head.fill.fgColor.rgb not in (None, "00000000"):
            c.fill = PatternFill("solid", fgColor=head.fill.fgColor.rgb)
    for i, w in enumerate([45, 45, 10, 40]):
        ws_s.column_dimensions[get_column_letter(cs + 1 + i)].width = w

    used = defaultdict(list)
    stats = {"g": 0, "y": 0, "r": 0}
    total = 0
    for r in range(hs + 1, ws_s.max_row + 1):
        v = ws_s.cell(r, cs).value
        if not v or not str(v).strip():
            continue
        total += 1
        lat = to_latin(v)
        t = tokens(v)
        ws_s.cell(r, cs + 1, lat)

        best, bs, note = None, 0, ""
        exact = by_key.get(" ".join(t), [])
        if exact:
            best, bs = exact[0], 100
            if len(exact) > 1:
                note = f"adliyada {len(exact)} ta bir xil F.I.Sh. (qatorlar: {', '.join(str(x['row']) for x in exact[:5])})"
        elif t:
            for a in by_block.get(coarse(t[0])[:3], []):
                s, n = score(t, a["t"])
                if s > bs:
                    best, bs, note = a, s, n

        c_name, c_pct, c_note = ws_s.cell(r, cs + 2), ws_s.cell(r, cs + 3), ws_s.cell(r, cs + 4)
        if best and bs >= MIN_ACCEPT:
            c_name.value, c_pct.value = best["name"], bs
            used[best["row"]].append(r)
            k = "g" if bs >= GREEN_FROM else "y"
        else:
            c_name.value, c_pct.value = None, bs or None
            if best and bs >= 85:
                note = f"taxminiy (olinmadi): {best['name']} ({bs}%) — qo'lda tekshiring"
            else:
                note = "adliyada topilmadi"
            k = "r"
        c_note.value = note or None
        c_name.fill = c_pct.fill = FILL[k]
        stats[k] += 1
        if total % 2000 == 0:
            print(f"  {total} ta qator...")

    # --- adliya varag'iga belgi ---
    mc = ws_a.max_column + 1
    ws_a.cell(ha, mc, "Soliqda bor (qator)").font = Font(bold=True)
    ws_a.column_dimensions[get_column_letter(mc)].width = 22
    for a in adl:
        rows = used.get(a["row"])
        c = ws_a.cell(a["row"], mc, ", ".join(map(str, rows)) if rows else None)
        c.fill = FILL["g"] if rows else FILL["r"]

    ws_s.auto_filter.ref = f"A{hs}:{get_column_letter(ws_s.max_column)}{ws_s.max_row}"
    ws_s.freeze_panes = ws_s.cell(hs + 1, 1)
    try:
        wb.save(out)
    except PermissionError:
        sys.exit(f"'{out}' ochiq turibdi — yopib, qayta ishga tushiring.")

    print(f"\nSaqlandi: {out}")
    print(f"  Soliqda jami: {total}")
    print(f"  Aniq topildi va olib o'tildi: {stats['g'] + stats['y']}")
    print(f"  Olib o'tilmadi (topilmadi yoki taxminiy): {stats['r']}")
    print(f"  Adliyadagi {len(adl)} tadan soliqqa mos kelgani: {len(used)}")


if __name__ == "__main__":
    main()
