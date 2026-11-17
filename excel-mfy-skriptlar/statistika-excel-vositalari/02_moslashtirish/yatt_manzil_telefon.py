"""
"1722 ЯТТ хоким ёрдамчиларига руйхат+.xlsx" dan MANZIL va TELEFON ni
F.I.Sh bo'yicha topib, "yatt o'tmagan.xlsx" ga olib o'tadi va faylni chiroyli bezaydi.

F.I.Sh qanday solishtiriladi:
  * kirill/lotin, O‘G‘LI / O'G'LI / OGLI / UGLI, QIZI / KIZI / KIZA, x=h, q=k,
    qo'sh harf, ortiqcha bo'shliq va belgilar tenglashtiriladi;
  * otasining ismi "-OVICH/-OVNA" yoki "... O'G'LI/QIZI" bo'lsa ham bir xil o'zak olinadi;
  * avval TUMAN ichida qidiriladi (bir xil ismli odamlar boshqa tumanlarda bo'lishi mumkin),
    bir tumanda bir necha mahallada bir xil ism bo'lsa - MAHALLA bo'yicha ajratiladi;
  * ishonchsiz moslikda ma'lumot YOZILMAYDI, "Tekshirish" varag'ida nomzodlar ko'rsatiladi.

Natija: yangi fayl (asl fayl o'zgarmaydi):
  * asosiy varaq: asl ustunlarga TEGILMAYDI; oxiriga yashil sarlavhali yangi ustunlar qo'shiladi:
    "Telefon (1722 dan)", "Manzil (1722 dan)", "MFY nomi", "Ko'cha", "Uy", "Tuman (1722 dan)", "F.I.Sh (1722 dagi)", "Moslik";
    sarlavha, chegara, kenglik, filtr, muzlatilgan sarlavha, chop etish sozlamalari bilan;
  * "Xulosa" varag'i - tumanlar kesimida natija;
  * "Tekshirish" varag'i - taxminiy / noaniq / topilmagan qatorlar.

O'rnatish:  pip install pandas openpyxl
"""
import os
import re
import sys
import traceback
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from difflib import SequenceMatcher

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))   # yoki r"Z:\МОЛИЯ БУЛИМИ\1"
SOURCE_FILE = "1722 ЯТТ хоким ёрдамчиларига руйхат+.xlsx"   # manzil/telefon olinadi
TARGET_FILE = "yatt o'tmagan.xlsx"                          # shu faylga yoziladi
OUTPUT_NAME = "yatt o'tmagan_natija.xlsx"

NORMALIZE_PHONES = True        # olib o'tilgan telefonlar "+998 97 123 45 67" ko'rinishida yoziladi
NEW_HEADERS = {                # yangi (olib o'tilgan) ustunlar sarlavhalari
    "phone": "Telefon (1722 dan)",
    "addr": "Manzil (1722 dan)",
    "tuman": "Tuman (1722 dan)",
    "mfy": "MFY nomi (1722 dan)",
    "street": "Ko'cha (1722 dan)",
    "house": "Uy (1722 dan)",
    "fio": "F.I.Sh (1722 dagi)",
    "status": "Moslik",
}
FUZZY_THRESHOLD = 0.92         # imlo farqi bilan moslash chegarasi (0..1)
FUZZY_MARGIN = 0.03
ALLOW_OTHER_TUMAN = True       # tumanda topilmasa, boshqa tumandan (faqat yagona bo'lsa) - sariq
MAX_JOIN = 3                   # bir odamning bir nechta manzili/telefoni bo'lsa, nechtasini qo'shish
# ======================================================

# ---------------------------------------------------------------- matn
CYR2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
    "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "",
    "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "ғ": "g",
    "қ": "k", "ҳ": "h",
}
APOS = re.compile(r"['`ʻʼ’‘´\"“”«»]")
PATRON_WORDS = {"ogli", "ugli", "oglu", "ugly", "kizi", "qizi", "kiza", "qiza", "kizy"}
PATRON_SUFFIX = re.compile(r"(ovich|evich|yevich|ivich|ovna|evna|yevna|ichna|ich|vna)$")
EMPTY_VALUES = {"", "0", "#н/д", "#n/a", "n/a", "nan", "none", "-", "—", "#н/а"}


def nfc(s):
    return unicodedata.normalize("NFC", str(s))


def isnan(v):
    try:
        return bool(pd.isna(v))
    except Exception:
        return False


def blank(v):
    return v is None or isnan(v) or str(v).strip().casefold() in EMPTY_VALUES


def lat(v):
    if blank(v):
        return ""
    s = " ".join(nfc(v).casefold().split())
    s = "".join(CYR2LAT.get(ch, ch) for ch in s)
    return APOS.sub("", s)


def red(tok):
    """Imlo variantlari: x=h, q=k, w=v, yo/yu/ya soddalashtiriladi, qo'sh harf -> bitta."""
    t = re.sub(r"[^a-z0-9]", "", tok)
    t = t.replace("x", "h").replace("q", "k").replace("w", "v")
    t = t.replace("dj", "j")      # Djumaev = Jumaev
    t = t.replace("ye", "e").replace("yo", "o").replace("yu", "u")
    return re.sub(r"(.)\1+", r"\1", t)


def name_keys(raw):
    """-> dict(full, stem, sf, s) yoki None.
    full: barcha so'zlar; stem: familiya+ism+ota o'zagi; sf: familiya+ism."""
    s = lat(raw)
    toks = [red(t) for t in re.findall(r"[a-z0-9]+", s)]
    toks = [t for t in toks if t]
    if not toks:
        return None
    core = [t for t in toks if t not in PATRON_WORDS and t not in ("ibn", "bin")]
    if not core:
        return None
    sur = re.sub(r"(ova|eva|yeva|iva)$", "#f", core[0])
    sur = re.sub(r"(ov|ev|yev|iv)$", "#", sur)        # Abdiev = Abdieva (o'zak)
    first = core[1] if len(core) > 1 else ""
    patr = core[2] if len(core) > 2 else ""
    # jinsi: familiya (-ova/-ov), otasining ismi (-ovna/-ovich), qizi/o'g'li
    g = ""
    if any(t in ("kizi", "qizi", "kiza", "qiza", "kizy") for t in toks) or re.search(r"(ovna|evna|ichna)$", patr):
        g = "f"
    elif any(t in ("ogli", "ugli", "oglu", "ugly") for t in toks) or re.search(r"(vich|ich)$", patr):
        g = "m"
    elif sur.endswith("#f"):
        g = "f"
    elif sur.endswith("#"):
        g = "m"
    patr = PATRON_SUFFIX.sub("", patr)
    patr = re.sub(r"(d|t)$", "t", patr)                # Soatmurod / Soatmurotovna
    return {
        "full": " ".join(toks),
        "stem": f"{sur}|{first}|{patr}",
        "sf": f"{sur}|{first}",
        "s": f"{sur} {first} {patr}".strip(),
        "g": g,
    }


def tuman_key(raw):
    s = lat(raw)
    city = bool(re.search(r"\b(shahri|shahar|sh|gorod)\b", s))
    toks = [t for t in re.findall(r"[a-z0-9]+", s)
            if t not in {"tuman", "tumani", "shahar", "shahri", "sh", "viloyati", "rayon", "gorod"}]
    k = red("".join(toks))
    return (k + "#shahar" if city and k else k), city


def mfy_key(raw):
    s = lat(raw)
    if not s or re.fullmatch(r"[\d\s]+", s):
        return ""
    s = re.sub(r"\bm\.?\s*f\.?\s*y\.?|\bmfy\b|\bmahallasi\b|\bmahalla\b|\bmaxallasi\b|\bmsg\b|\bqfy\b", " ", s)
    return red(s)


def clean_phone(v):
    if blank(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    raw = str(v).strip()
    d = re.sub(r"\D", "", raw)
    if NORMALIZE_PHONES:
        if len(d) == 9:
            d = "998" + d
        if len(d) == 12 and d.startswith("998"):
            return f"+998 {d[3:5]} {d[5:8]} {d[8:10]} {d[10:12]}"
    return raw if len(d) >= 7 else ""


def clean_text(v):
    if blank(v):
        return ""
    return " ".join(str(v).split())


# ------------------------------------------------- manzildan ko'cha va uy
_Q = r"[\"'`ʻʼ’‘´“”«»]"
_SUF = (r"(?:ko{q}?cha(?:si|sida)?|kuchasi|кўчаси|кучаси|кўчасида|kocha|"
        r"tor\s+ko{q}?chasi|berk\s+ko{q}?chasi|туп\.?|тупиги|проезд)").format(q=_Q)
STREET_AFTER = re.compile(rf"([^,;:]+?)\s+{_SUF}(?=[\s,.;:]|$)", re.I)
STREET_BEFORE = re.compile(r"(?:\bул\.|\bулица|\bko{q}?ch\.)\s*([^,;:]+)".format(q=_Q), re.I)
STREET_STOP = re.compile(
    r"^.*\b(?:mahallasi(?:da)?|maxallasi|mahalla|mfy|m\.f\.y\.?|qfy|msg|shfy|"
    r"маҳалласи|махалласи|мфй|мсг|қфй|кфй|туман[и]?|tumani|viloyati|вилояти|"
    r"шаҳри|шахри|shahri|qo{q}?rg{q}?oni|қўрғони|курғони)\s+".format(q=_Q), re.I)
HOUSE = re.compile(
    r"(?:\b(?:дом|uy|уй|д\.)\s*[:.]?\s*([0-9][0-9a-zа-яўқғҳ/\-]*))|"
    r"(?:\b([0-9][0-9a-zа-яўқғҳ/\-]*)\s*-?\s*(?:uy|уй|uyda|уйда)\b)", re.I)


def extract_street(addr):
    """Manzil matnidan ko'cha nomini ajratadi, topilmasa ''."""
    if blank(addr):
        return ""
    text = " ".join(nfc(addr).split())
    m = STREET_AFTER.search(text)
    name = m.group(1) if m else ""
    if not name:
        m = STREET_BEFORE.search(text)
        name = m.group(1) if m else ""
        name = re.split(r"\s+(?:дом|д\.|uy|уй)\b", name, flags=re.I)[0]
    name = STREET_STOP.sub("", name)
    name = re.sub(r"[\"“”«»]", "", name)
    name = re.sub(r"^\s*(?:улица|ул\.?)\s*", "", name, flags=re.I)
    name = name.strip(" .-")
    if not name or len(name) < 2 or re.fullmatch(r"[\d\s]+", name) or len(name) > 40:
        return ""
    return name


def extract_house(addr):
    if blank(addr):
        return ""
    for m in HOUSE.finditer(nfc(addr)):
        v = (m.group(1) or m.group(2) or "").strip("-/ ")
        if v and v not in ("0", "00"):
            return v
    return ""


def join_unique(vals, limit=MAX_JOIN):
    out = []
    for v in vals:
        if v and v not in out:
            out.append(v)
    return " | ".join(out if limit is None else out[:limit])


# ----------------------------------------------------------- manba fayl
def detect_cols(values, want):
    """values: sarlavha qatori. want: {kalit: [regex,...]}"""
    found = {}
    for j, v in enumerate(values):
        h = lat(v)
        if not h:
            continue
        for k, pats in want.items():
            if k not in found and any(re.search(p, h) for p in pats):
                found[k] = j
                break
    return found


SRC_WANT = {
    "fio": [r"f\.?\s*i\.?\s*sh", r"\bfio\b", r"familiya", r"fish\b"],
    "phone": [r"telefon", r"\btel\b", r"phone"],
    "addr": [r"manzil", r"adres", r"address"],
    "tuman": [r"tuman", r"shahar", r"rayon"],
    "mfy": [r"mfy\s*nom", r"mahalla", r"\bmfy\b(?!.*stir)"],
}


def load_source(path):
    xl = pd.ExcelFile(path)
    rows = []
    for name in xl.sheet_names:
        head = pd.read_excel(xl, sheet_name=name, header=None, nrows=15, dtype=str)
        hdr = cols = None
        for i in range(len(head)):
            c = detect_cols(head.iloc[i].tolist(), SRC_WANT)
            if "fio" in c and ("phone" in c or "addr" in c):
                hdr, cols = i, c
                break
        if cols is None:
            print(f"  [{name}] F.I.Sh / manzil / telefon ustunlari yo'q - o'tkazildi")
            continue
        df = pd.read_excel(xl, sheet_name=name, header=hdr, dtype=object)
        get = lambda k: df.iloc[:, cols[k]].tolist() if k in cols else [None] * len(df)
        n0 = len(rows)
        for fio, ph, ad, tu, mf in zip(get("fio"), get("phone"), get("addr"),
                                       get("tuman"), get("mfy")):
            k = name_keys(fio)
            if not k:
                continue
            tk, city = tuman_key(tu)
            rows.append({"fio": clean_text(fio), "k": k, "tk": tk, "city": city,
                         "mfy": mfy_key(mf), "mfy_raw": clean_text(mf),
                         "phone": clean_phone(ph), "addr": clean_text(ad),
                         "tuman_raw": clean_text(tu)})
        print(f"  [{name}] ustunlar: " + ", ".join(f"{k}={df.columns[v]}" for k, v in cols.items())
              + f" | qatorlar: {len(rows) - n0}")
    return rows


class Index:
    def __init__(self, rows):
        self.rows = rows
        self.by = {lvl: defaultdict(list) for lvl in ("full", "stem", "sf")}
        self.by_tuman = defaultdict(list)
        for i, r in enumerate(rows):
            for lvl in self.by:
                self.by[lvl][(r["tk"], r["k"][lvl])].append(i)
                self.by[lvl][("*", r["k"][lvl])].append(i)
            self.by_tuman[r["tk"]].append(i)
        self.tumans = list(self.by_tuman)
        self.cache = {}

    def resolve_tuman(self, tk):
        if not tk:
            return None
        if tk in self.by_tuman:
            return tk
        best = max(((SequenceMatcher(None, tk, t).ratio(), t) for t in self.tumans if t),
                   default=(0, None))
        return best[1] if best[0] >= 0.8 else None

    def decide(self, idxs, q_mfy):
        """Bir xil ismli yozuvlar: bitta odammi yoki bir necha?"""
        by_t = defaultdict(lambda: defaultdict(list))
        for i in idxs:
            by_t[self.rows[i]["tk"]][self.rows[i]["mfy"] or "?"].append(i)
        groups = []                      # (mahalla kalitlari, yozuvlar) - har biri bitta odam
        for g in by_t.values():
            named = [m for m in g if m != "?"]
            if len(named) <= 1:          # mahallasiz yozuvlar yagona mahalladagisiga qo'shiladi
                groups.append((set(named), [i for lst in g.values() for i in lst]))
            else:
                groups += [({m}, g[m]) for m in named]
                if "?" in g:
                    groups.append((set(), g["?"]))
        if len(groups) == 1:
            return "ok", groups[0][1]
        if q_mfy:
            hit = [gr for gr in groups if any(
                m == q_mfy or SequenceMatcher(None, m, q_mfy).ratio() >= 0.85 for m in gr[0])]
            if len(hit) == 1:
                return "ok_mfy", hit[0][1]
        return "amb", idxs

    def lookup(self, fio, tuman_raw, mfy_raw):
        key = (fio, tuman_raw, mfy_raw)
        if key not in self.cache:
            self.cache[key] = self._lookup(fio, tuman_raw, mfy_raw)
        return self.cache[key]

    def _lookup(self, fio, tuman_raw, mfy_raw):
        """-> (yozuv_indekslari, daraja, izoh)"""
        k = name_keys(fio)
        if not k:
            return [], "bosh", "F.I.Sh bo'sh"
        tk = self.resolve_tuman(tuman_key(tuman_raw)[0])
        q_mfy = mfy_key(mfy_raw)
        scopes = [(tk, "tuman ichida")] if tk else []
        if ALLOW_OTHER_TUMAN or not tk:
            scopes.append(("*", "boshqa tumanda" if tk else "tuman topilmadi"))

        amb = None
        for scope, where in scopes:
            for lvl, note in (("full", "aniq"), ("stem", "aniq (yozuv farqi)"),
                              ("sf", "familiya+ism (otasining ismi yo'q/farq)")):
                if lvl == "sf" and k["stem"].split("|")[2] and scope == "*":
                    continue
                idxs = self.by[lvl].get((scope, k[lvl]))
                if not idxs:
                    continue
                if lvl == "sf":   # otasining ismi ikkala tomonda ham bor va farq qilsa - boshqa odam
                    qp = k["stem"].split("|")[2]
                    idxs = [i for i in idxs if not qp or not self.rows[i]["k"]["stem"].split("|")[2]]
                    if not idxs:
                        continue
                st, sel = self.decide(idxs, q_mfy)
                if st == "amb":
                    amb = amb or (idxs, f"noaniq: {len(idxs)} ta bir xil ismli ({where}), "
                                        "mahalla bo'yicha ajratib bo'lmadi")
                    continue
                exact = lvl in ("full", "stem") and scope != "*"
                if exact:
                    extra = " + mahalla bo'yicha tanlandi" if st == "ok_mfy" else ""
                    return sel, "aniq", note + extra
                return sel, "taxminiy", f"{note}, {where}"
            if amb:
                return amb[0], "noaniq", amb[1]
            # imlo farqi (fuzzy) - faqat tuman ichida
            if scope != "*" and len(k["s"]) >= 8:
                sm = SequenceMatcher(None)
                sm.set_seq2(k["s"])
                scored = []
                sur = k["stem"].split("|")[0][:3]
                for i in self.by_tuman[scope]:
                    rk = self.rows[i]["k"]
                    rs = rk["s"]
                    if rs[:2] != sur[:2] or (k["g"] and rk["g"] and k["g"] != rk["g"]):
                        continue
                    sm.set_seq1(rs)
                    if sm.real_quick_ratio() < FUZZY_THRESHOLD or sm.quick_ratio() < FUZZY_THRESHOLD:
                        continue
                    r = sm.ratio()
                    if r >= FUZZY_THRESHOLD:
                        scored.append((r, i))
                if scored:
                    scored.sort(reverse=True)
                    r1, i1 = scored[0]
                    best = [i for r, i in scored if self.rows[i]["k"]["s"] == self.rows[i1]["k"]["s"]]
                    other = next((r for r, i in scored
                                  if self.rows[i]["k"]["s"] != self.rows[i1]["k"]["s"]), None)
                    if other is not None and r1 - other < FUZZY_MARGIN:
                        return [i for _, i in scored], "noaniq", "noaniq: bir necha o'xshash ism"
                    st, sel = self.decide(best, q_mfy)
                    if st != "amb":
                        return sel, "taxminiy", f"imlo farqi ({r1:.0%}), tuman ichida"
                    return best, "noaniq", "noaniq: o'xshash ism bir necha mahallada"
        return [], "topilmadi", "topilmadi"


# -------------------------------------------------------------- bezak
C_HEAD = "1F4E78"
C_BORDER = "BFBFBF"
FILL = {
    "aniq": PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA"),
    "taxminiy": PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC"),
    "noaniq": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
    "topilmadi": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
}
LABEL = {"aniq": "✔ Topildi", "taxminiy": "≈ Taxminiy", "noaniq": "? Noaniq",
         "topilmadi": "✘ Topilmadi", "bosh": "—"}
THIN = Side(style="thin", color=C_BORDER)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", start_color=C_HEAD, end_color=C_HEAD)
C_NEW_HEAD = "375623"          # olib o'tilgan ustunlar sarlavhasi - yashil
NEW_HEAD_FILL = PatternFill("solid", start_color=C_NEW_HEAD, end_color=C_NEW_HEAD)
NEW_FILL = PatternFill("solid", start_color="EEF6E8", end_color="EEF6E8")
BODY_FONT = Font(name="Calibri", size=10.5)


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.font, cell.fill, cell.border = HEAD_FONT, HEAD_FILL, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32


def finish_sheet(ws, hdr, ncol, widths, landscape=True):
    ws.freeze_panes = ws.cell(hdr + 1, 2 if ncol > 3 else 1)
    ws.auto_filter.ref = f"A{hdr}:{get_column_letter(ncol)}{max(ws.max_row, hdr + 1)}"
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_view.zoomScale = 100
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hdr}:{hdr}"
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4


def auto_width(ws, col, hdr, lo=8, hi=60):
    best = lo
    for r in range(hdr, min(ws.max_row, hdr + 3000) + 1):
        v = ws.cell(r, col).value
        if v is not None:
            best = max(best, min(hi, int(max(len(x) for x in str(v).split("\n")) * 1.1) + 3))
    return best


# ---------------------------------------------------------------- asosiy
TGT_WANT = {
    "fio": [r"f\.?\s*i\.?\s*sh", r"\bfio\b", r"familiya", r"fish\b"],
    "phone": [r"telefon", r"\btel\b", r"phone"],
    "addr": [r"manzil", r"adres"],
    "tuman": [r"tuman", r"shahar", r"rayon"],
    "mfy": [r"mahalla", r"\bmfy\b"],
    "holat": [r"holat"],
    "viloyat": [r"viloyat"],
}


def find_file(name, title):
    files = sorted(f for f in os.listdir(BASE_DIR)
                   if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"))
    key = re.sub(r"[^a-z0-9]", "", lat(name))
    for f in files:
        if re.sub(r"[^a-z0-9]", "", lat(f)) == key:
            return os.path.join(BASE_DIR, f)
    print(f"'{name}' topilmadi. {title}")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    try:
        return os.path.join(BASE_DIR, files[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else find_file(SOURCE_FILE, "Manba (1722 YaTT) faylni tanlang:")
    tgt = sys.argv[2] if len(sys.argv) > 2 else find_file(TARGET_FILE, "YaTT o'tmagan faylni tanlang:")
    if not src or not tgt:
        print("Fayllar tanlanmadi.")
        return
    print("Manba :", os.path.basename(src))
    print("Natija:", os.path.basename(tgt))

    print("Manba fayl o'qilmoqda...")
    rows = load_source(src)
    if not rows:
        print("Manba faylda F.I.Sh ustuni topilmadi.")
        return
    idx = Index(rows)
    print(f"Manba: {len(rows)} ta yozuv, {len(idx.tumans)} ta tuman")

    wb = load_workbook(tgt)
    ws = None
    for w in wb.worksheets:
        for r in range(1, 11):
            c = detect_cols([x.value for x in w[r]], TGT_WANT)
            if "fio" in c:
                ws, hdr, cols = w, r, c
                break
        if ws:
            break
    if ws is None:
        print("Natija faylida F.I.Sh ustuni topilmadi.")
        return
    print(f"Varaq: [{ws.title}] | sarlavha qatori: {hdr} | ustunlar: "
          + ", ".join(f"{k}={get_column_letter(v + 1)}" for k, v in cols.items()))

    orig_last = max(c.column for c in ws[hdr] if c.value not in (None, ""))
    c_fio = cols["fio"] + 1
    c_tum = cols["tuman"] + 1 if "tuman" in cols else None
    c_mfy = cols["mfy"] + 1 if "mfy" in cols else None
    # Olib o'tilgan ma'lumotlar ALOHIDA ustunlarga yoziladi - asl ustunlarga tegilmaydi
    c_ph, c_ad, c_mf, c_ko, c_uy, c_tu, c_src, c_st = (orig_last + i for i in range(1, 9))
    for c, title in ((c_ph, NEW_HEADERS["phone"]), (c_ad, NEW_HEADERS["addr"]),
                     (c_mf, NEW_HEADERS["mfy"]),
                     (c_ko, NEW_HEADERS["street"]), (c_uy, NEW_HEADERS["house"]),
                     (c_tu, NEW_HEADERS["tuman"]), (c_src, NEW_HEADERS["fio"]),
                     (c_st, NEW_HEADERS["status"])):
        ws.cell(hdr, c).value = title
    last = c_st

    stats = Counter()
    per_tuman = defaultdict(Counter)
    review = []
    for r in range(hdr + 1, ws.max_row + 1):
        fio = ws.cell(r, c_fio).value
        if blank(fio):
            continue
        tum = ws.cell(r, c_tum).value if c_tum else None
        mfy = ws.cell(r, c_mfy).value if c_mfy else None
        sel, level, note = idx.lookup(str(fio), tum, mfy)
        stats[level] += 1
        per_tuman[clean_text(tum) or "(tuman yo'q)"][level] += 1
        if level in ("aniq", "taxminiy") and sel:
            ws.cell(r, c_ph).value = join_unique(rows[i]["phone"] for i in sel) or None
            ws.cell(r, c_ad).value = join_unique((rows[i]["addr"] for i in sel), limit=None) or None
            ws.cell(r, c_mf).value = join_unique(rows[i]["mfy_raw"] for i in sel) or None
            ws.cell(r, c_ko).value = join_unique(extract_street(rows[i]["addr"]) for i in sel) or None
            ws.cell(r, c_uy).value = join_unique(extract_house(rows[i]["addr"]) for i in sel) or None
            ws.cell(r, c_tu).value = join_unique(rows[i]["tuman_raw"] for i in sel) or None
            ws.cell(r, c_src).value = join_unique(rows[i]["fio"] for i in sel)
        cell_st = ws.cell(r, c_st)
        cell_st.value = LABEL.get(level, level)
        cell_st.fill = FILL.get(level, PatternFill())
        if level != "aniq":
            cands = "\n".join(
                f'{rows[i]["fio"]} | {rows[i]["tuman_raw"]} | {rows[i]["mfy_raw"]} | '
                f'{rows[i]["phone"] or "-"} | {rows[i]["addr"] or "-"}' for i in sel[:6])
            review.append([r, clean_text(fio), clean_text(tum), clean_text(mfy),
                           LABEL.get(level, level), note, cands])

    # ---------- asosiy varaq bezagi
    ncol = last
    style_header(ws, hdr, ncol)
    new_cols = (c_ph, c_ad, c_mf, c_ko, c_uy, c_tu, c_src, c_st)
    for c in new_cols:                       # yangi ustunlar sarlavhasi - boshqa rang
        ws.cell(hdr, c).fill = NEW_HEAD_FILL
    sep = Side(style="medium", color=C_NEW_HEAD)
    for r in range(hdr, ws.max_row + 1):
        if r > hdr:
            ws.row_dimensions[r].height = None
        for c in range(1, ncol + 1):
            cell = ws.cell(r, c)
            if r > hdr:
                if c != c_st:
                    cell.font = BODY_FONT
                cell.alignment = Alignment(vertical="center",
                                           horizontal="center" if c in (c_ph, c_st, c_uy) else "left",
                                           wrap_text=(c in (c_ad, c_mf)))
                if c in (c_ph, c_ad, c_mf, c_ko, c_uy, c_tu, c_src):
                    cell.fill = NEW_FILL
            cell.border = Border(left=sep if c == c_ph else THIN, right=THIN, top=THIN, bottom=THIN)
        if r > hdr:
            ws.cell(r, c_ph).number_format = "@"
            ws.cell(r, c_uy).number_format = "@"
            ws.cell(r, c_st).font = Font(name="Calibri", size=10.5, bold=True)
    widths = []
    for c in range(1, ncol + 1):
        if c == c_ad:
            widths.append(55)
        elif c == c_ph:
            widths.append(21)
        elif c == c_st:
            widths.append(14)
        elif c == c_tu:
            widths.append(20)
        elif c == c_mf:
            widths.append(24)
        elif c == c_ko:
            widths.append(24)
        elif c == c_uy:
            widths.append(10)
        elif c == c_src:
            widths.append(38)
        else:
            widths.append(auto_width(ws, c, hdr, 10, 48))
    finish_sheet(ws, hdr, ncol, widths)
    # o'ralgan matn to'liq ko'rinishi uchun qator balandligi (Excel o'zi hisoblamaydi)
    for r in range(hdr + 1, ws.max_row + 1):
        lines = 1
        for c in (c_ad, c_mf):
            v = ws.cell(r, c).value
            if v:
                per = max(1, int(widths[c - 1] * 1.15))
                lines = max(lines, sum(-(-len(part) // per) or 1 for part in str(v).split("\n")))
        ws.row_dimensions[r].height = 15 * lines if lines > 1 else None
    # zebra - faqat asl ustunlarda (filtrlanganda ham to'g'ri ishlaydi)
    fl = get_column_letter(c_fio)
    ws.conditional_formatting.add(
        f"A{hdr + 1}:{get_column_letter(orig_last)}{ws.max_row}", FormulaRule(
            formula=[f"AND(MOD(SUBTOTAL(3,${fl}${hdr + 1}:${fl}{hdr + 1}),2)=0,"
                     f"${fl}{hdr + 1}<>\"\")"],
            fill=PatternFill("solid", start_color="F3F6FA", end_color="F3F6FA"), stopIfTrue=False))

    # ---------- Xulosa
    for name in ("Xulosa", "Tekshirish"):
        if name in wb.sheetnames:
            del wb[name]
    sx = wb.create_sheet("Xulosa", 0)
    sx["A1"] = "YaTT o'tmaganlar: manzil va telefon olib o'tish natijasi"
    sx["A1"].font = Font(name="Calibri", size=14, bold=True, color=C_HEAD)
    sx["A2"] = (f"Manba: {os.path.basename(src)}   •   "
                f"Sana: {datetime.now():%d.%m.%Y %H:%M}")
    sx["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")
    head = ["Tuman", "Jami", "✔ Topildi", "≈ Taxminiy", "? Noaniq", "✘ Topilmadi", "Topildi, %"]
    sx.append([])
    sx.append(head)
    h = sx.max_row
    style_header(sx, h, len(head))
    for t in sorted(per_tuman):
        c = per_tuman[t]
        tot = sum(c.values())
        sx.append([t, tot, c["aniq"], c["taxminiy"], c["noaniq"], c["topilmadi"], None])
        sx.cell(sx.max_row, 7).value = f"=IF(B{sx.max_row}=0,0,(C{sx.max_row}+D{sx.max_row})/B{sx.max_row})"
    first, lastr = h + 1, sx.max_row
    sx.append(["JAMI"] + [f"=SUM({get_column_letter(c)}{first}:{get_column_letter(c)}{lastr})"
                          for c in range(2, 7)] +
              [f"=IF(B{lastr + 1}=0,0,(C{lastr + 1}+D{lastr + 1})/B{lastr + 1})"])
    for r in range(first, sx.max_row + 1):
        for c in range(1, len(head) + 1):
            cell = sx.cell(r, c)
            cell.border = BOX
            cell.font = Font(name="Calibri", size=11, bold=(r == sx.max_row))
            cell.alignment = Alignment(horizontal="left" if c == 1 else "center", vertical="center")
            if c >= 2:
                cell.number_format = "0.0%" if c == 7 else "#,##0"
        if r == sx.max_row:
            for c in range(1, len(head) + 1):
                sx.cell(r, c).fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
    for c, key in ((3, "aniq"), (4, "taxminiy"), (5, "noaniq"), (6, "topilmadi")):
        sx.cell(h, c).fill = PatternFill("solid", start_color=C_HEAD, end_color=C_HEAD)
        for r in range(first, lastr + 1):
            sx.cell(r, c).fill = FILL[key]
    for i, w in enumerate([28, 10, 12, 12, 11, 13, 12], 1):
        sx.column_dimensions[get_column_letter(i)].width = w
    sx.freeze_panes = f"A{h + 1}"
    sx.page_setup.orientation = "portrait"
    sx.page_setup.paperSize = sx.PAPERSIZE_A4
    sx.page_setup.fitToWidth, sx.page_setup.fitToHeight = 1, 0
    sx.sheet_properties.pageSetUpPr.fitToPage = True
    sx.sheet_view.showGridLines = False
    note_r = sx.max_row + 2
    for i, line in enumerate([
        "Ranglar:",
        "✔ Topildi — F.I.Sh va tuman to'liq mos.",
        "≈ Taxminiy — yozuvda kichik farq yoki boshqa tumanda topildi: tekshirib chiqing.",
        "? Noaniq — bir xil ismli bir nechta odam: ma'lumot yozilmadi, \"Tekshirish\" varag'ida nomzodlar bor.",
        "✘ Topilmadi — manba faylda bu odam yo'q.",
    ]):
        sx.cell(note_r + i, 1).value = line
        sx.cell(note_r + i, 1).font = Font(name="Calibri", size=10, bold=(i == 0), color="404040")

    # ---------- Tekshirish
    sr = wb.create_sheet("Tekshirish")
    head = ["Qator №", "F.I.Sh", "Tuman", "Mahalla", "Moslik", "Izoh",
            "Nomzodlar (F.I.Sh | tuman | MFY | telefon | manzil)"]
    sr.append(head)
    style_header(sr, 1, len(head))
    order = {"≈ Taxminiy": 0, "? Noaniq": 1, "✘ Topilmadi": 2}
    review.sort(key=lambda x: (order.get(x[4], 3), x[2], x[1]))
    rev_fill = {"≈ Taxminiy": FILL["taxminiy"], "? Noaniq": FILL["noaniq"], "✘ Topilmadi": FILL["topilmadi"]}
    for row in review:
        sr.append(row)
        r = sr.max_row
        sr.cell(r, 1).hyperlink = f"#'{ws.title}'!A{row[0]}"
        sr.cell(r, 1).font = Font(name="Calibri", size=10.5, color="0563C1", underline="single")
        for c in range(1, len(head) + 1):
            cell = sr.cell(r, c)
            cell.border = BOX
            if c != 1:
                cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=c in (6, 7),
                                       horizontal="center" if c in (1, 5) else "left")
        sr.cell(r, 5).fill = rev_fill.get(row[4], PatternFill())
    finish_sheet(sr, 1, len(head), [9, 36, 18, 20, 13, 40, 90])

    wb.active = wb.sheetnames.index(ws.title)

    out = os.path.join(BASE_DIR, OUTPUT_NAME)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"yatt_natija_{datetime.now():%H%M%S}.xlsx")
        print("Natija fayli Excelda ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)

    tot = sum(v for k, v in stats.items() if k != "bosh")
    ok = stats["aniq"] + stats["taxminiy"]
    print(f"Jami: {tot} | topildi: {stats['aniq']} | taxminiy: {stats['taxminiy']} | "
          f"noaniq: {stats['noaniq']} | topilmadi: {stats['topilmadi']} | "
          f"qamrov: {ok / tot:.1%}" if tot else "Qator yo'q")
    print("Saqlandi:", out)
    print("  Xulosa     - tumanlar kesimida natija")
    print("  Tekshirish - taxminiy / noaniq / topilmagan qatorlar (nomzodlar bilan)")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
