"""
"SSP Jami a'zo ..." jadvaliga korxona INN sini "110.xlsx" dan nomi bo'yicha olib o'tadi.

Nomlar solishtirilishi (har bir korxona uchun):
  1. Asosiy nom (qo'shtirnoq ichidagi) + tashkiliy shakl (MChJ, XK, OK, F/X, ...)
     ikkala faylda bir xil yozuvga keltiriladi: kirill/lotin, tutuq belgilari,
     "mas`uliyati cheklangan jamiyati" = "MCHJ", "xususiy korxonasi" = "XK" va h.k.
  2. Aniq mos  ->  imlo farqi (q/k, x/h, qo'sh harf)  ->  so'z tartibi farqi  ->  o'xshashlik (fuzzy)
  3. Tashkiliy shakl mos kelmasa yoki bir nomga bir necha INN to'g'ri kelsa,
     noto'g'ri INN yozmaslik uchun YOZILMAYDI va hisobotga tushadi.

Natija: asl formatlar saqlanadi (faqat INN ustuni qo'shiladi/to'ldiriladi),
oxirida "INN_hisobot" varag'i qo'shiladi.

O'rnatish:  pip install pandas openpyxl
"""
import os
import re
import shutil
import sys
import traceback
import unicodedata
from collections import Counter
from copy import copy
from datetime import datetime
from difflib import SequenceMatcher

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # yoki r"Z:\МОЛИЯ БУЛИМИ\1"
TARGET_FILE = "SSP Jami a'zo 221 904.xlsx"   # INN yoziladigan fayl
REF_FILE = "110.xlsx"                        # INN olinadigan fayl
REF_SHEETS = ["Лист2"]          # 110 fayldan faqat shu varaq(lar); [] = hamma varaq
OUTPUT_NAME = "natija_SSP_INN.xlsx"
IN_PLACE = False               # True: INN shu faylning o'ziga yoziladi (zaxira nusxa bilan)
HEADER_INN = "INN"             # yangi ustun sarlavhasi (ustun yo'q bo'lsa)
HUDUD_FILTER = "surxondaryo"   # bo'sh emas bo'lsa: faqat "Hudud" ustunida shu so'z bor qatorlar
ONLY_VISIBLE_ROWS = False      # True: filtr yashirgan qatorlar o'tkazib yuboriladi
FUZZY_THRESHOLD = 0.90         # imlo farqi bilan moslash chegarasi (0..1)
FUZZY_MARGIN = 0.03            # eng yaxshi nomzod ikkinchisidan shuncha ustun bo'lsin
ALLOW_FORM_MISMATCH = False    # True: nom bir xil, lekin shakl boshqa bo'lsa ham yozish (sariq)
OVERWRITE_EXISTING = False     # False: to'ldirilgan INN kataklariga tegilmaydi
INN_AS_TEXT = True
HIGHLIGHT = True               # sariq = taxminiy, to'q sariq = noaniq/topilmadi
ADD_REPORT_SHEETS = True
VERIFY = True                  # saqlagach, asl fayl bilan solishtirish (katta faylda sekinroq)
SKIP_SHEETS = ["innhisobot"]
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
APOS = re.compile(r"['`ʻʼ’‘´]")
QUOTES = "\"“”«»„‟"
PATRONYMIC = {"ogli", "ugli", "oglu", "uglu", "kizi", "qizi"}

FORM_PATTERNS = [   # apostrofsiz lotin matnga
    ("MCHJ", r"\bmchj\b|\bmasuliyati\s+cheklangan\s+jamiyat\w*"),
    ("XORIJ", r"\bxorijiy\s+korxona\w*"),
    ("XK", r"\bxususiy\s+korxona\w*|\bxk\b"),
    ("OK", r"\boilaviy\s+korxona\w*|\bok\b"),
    ("FX", r"\bfermer\s+xo\w*jalig\w*|\bf\s*/\s*x\b|\bfx\b"),
    ("AJ", r"\baksiyadorlik\s+jamiyat\w*|\baj\b"),
    ("QK", r"\bqoshma\s+korxona\w*|\bqk\b"),
    ("DK", r"\bdavlat\s+korxona\w*|\bdk\b"),
    ("YATT", r"\byakka\s+tartibdagi\s+tadbirkor\w*|\byatt\b|\bytt\b"),
]


def nfc(s):
    return unicodedata.normalize("NFC", str(s))


def _isnan(v):
    try:
        return bool(pd.isna(v))
    except Exception:
        return False


def lat(v):
    """Kichik harf, lotinga o'tkazilgan, bo'shliqlar bir xil (tinish belgilari saqlanadi)."""
    if v is None or _isnan(v):
        return ""
    s = " ".join(nfc(v).casefold().split())
    return "".join(CYR2LAT.get(ch, ch) for ch in s)


def reduce_key(s):
    """Imlo variantlarini tenglashtiradi: x=h, q=k, w=v, qo'sh harf -> bitta."""
    s = re.sub(r"[^a-z0-9]", "", s)
    s = s.replace("x", "h").replace("q", "k").replace("w", "v")
    return re.sub(r"(.)\1+", r"\1", s)


def fkey(v):
    return re.sub(r"[^a-z0-9]", "", APOS.sub("", lat(v)))


def parse_name(raw):
    """-> (asosiy_nom, tashkiliy_shakllar). Qo'shtirnoq ichi asosiy nom hisoblanadi."""
    s = lat(raw)
    idx = [i for i, ch in enumerate(s) if ch in QUOTES]
    if len(idx) >= 2:
        core = s[idx[0] + 1: idx[-1]]
        rest = s[:idx[0]] + " " + s[idx[-1] + 1:]
    else:
        core = s.replace("\"", " ")
        rest = core
    rest = APOS.sub("", rest)
    forms = {code for code, pat in FORM_PATTERNS if re.search(pat, rest)}
    if len(idx) < 2:   # qo'shtirnoqsiz: shakl so'zlarini asosiy nomdan olib tashlaymiz
        c = APOS.sub("", core)
        for _, pat in FORM_PATTERNS:
            c = re.sub(pat, " ", c)
        core = c
    return APOS.sub("", core), forms


def keys_of(core):
    toks = re.findall(r"[a-z0-9]+", core)
    strict = "".join(toks)
    rtoks = [reduce_key(t) for t in toks]
    rtoks = [t for t in rtoks if t]
    relaxed = "".join(rtoks)
    pat = {reduce_key(p) for p in PATRONYMIC}
    sorted_key = "".join(sorted(t for t in rtoks if t not in pat))
    return strict, relaxed, sorted_key, rtoks


def form_state(q, r):
    """'same' | 'unknown' (bir tomonda shakl yo'q) | 'diff'"""
    if not q and not r:
        return "same"
    if not q or not r:
        return "unknown"
    return "same" if (q & r) else "diff"


def clean_inn(v):
    if v is None or _isnan(v):
        return ""
    s = re.sub(r"\.0+$", "", str(v).strip())
    return s if s.isdigit() and 9 <= len(s) <= 14 else ""


# ------------------------------------------------- INN manbasi (110.xlsx)
class Ref:
    def __init__(self):
        self.items = []                     # (relaxed, forms, inn, nom)
        self.strict, self.relaxed, self.sorted, self.tok = {}, {}, {}, {}
        self.cache = {}

    def add(self, name, inn):
        core, forms = parse_name(name)
        strict, relaxed, skey, rtoks = keys_of(core)
        if not strict or not inn:
            return
        i = len(self.items)
        self.items.append((relaxed, forms, inn, str(name).strip()))
        self.strict.setdefault(strict, []).append(i)
        self.relaxed.setdefault(relaxed, []).append(i)
        if len(rtoks) >= 2:
            self.sorted.setdefault(skey, []).append(i)
        for t in set(rtoks):
            if len(t) >= 3:
                self.tok.setdefault(t, []).append(i)

    # ---- bir nechta nomzoddan tanlash
    def pick(self, idxs, qforms):
        groups = {"same": [], "unknown": [], "diff": []}
        for i in idxs:
            groups[form_state(qforms, self.items[i][1])].append(i)
        for quality in ("same", "unknown"):
            lst = groups[quality]
            if lst:
                inns = {self.items[i][2] for i in lst}
                if len(inns) == 1:
                    return "ok", lst[0], quality, []
                return "amb", lst[0], quality, self.cands(lst)
        if groups["diff"]:
            inns = {self.items[i][2] for i in groups["diff"]}
            return (("form" if len(inns) == 1 else "amb"), groups["diff"][0], "diff",
                    self.cands(groups["diff"]))
        return None, None, None, []

    def cands(self, idxs):
        seen, out = set(), []
        for i in idxs:
            inn, nom = self.items[i][2], self.items[i][3]
            if inn not in seen:
                seen.add(inn)
                out.append(f"{inn} - {nom}")
        return out

    def lookup(self, name):
        """-> (inn|None, holat, topilgan_nom, o'xshashlik, daraja, nomzodlar)"""
        if name not in self.cache:
            self.cache[name] = self._lookup(name)
        return self.cache[name]

    def _lookup(self, name):
        core, qforms = parse_name(name)
        strict, relaxed, skey, rtoks = keys_of(core)
        if not strict:
            return None, "nom bo'sh", "", 0.0, "bosh", []

        amb = formmis = None
        stages = [("aniq", self.strict, strict, 1.0),
                  ("aniq (imlo farqi)", self.relaxed, relaxed, 1.0)]
        if len(rtoks) >= 2:
            stages.append(("so'z tartibi farq qiladi", self.sorted, skey, 0.98))
        for label, store, key, score in stages:
            idxs = store.get(key)
            if not idxs:
                continue
            st, i, quality, cl = self.pick(idxs, qforms)
            item = self.items[i] if i is not None else None
            if st == "ok":
                if label.startswith("so'z"):
                    return item[2], f"taxminiy ({label})", item[3], score, "taxminiy", []
                if quality == "unknown":
                    return (item[2], f"taxminiy ({label}, tashkiliy shakl ko'rsatilmagan)",
                            item[3], score, "taxminiy", [])
                return item[2], label, item[3], score, "aniq", []
            if st == "amb":
                amb = amb or (item[3], score, cl)
            elif st == "form":
                formmis = formmis or (item, score, cl)

        # ---- fuzzy (nomzodlar so'zlar bo'yicha tanlanadi)
        best_hint = (0.0, "")
        if len(relaxed) >= 5:
            cnt = Counter()
            # kam uchraydigan so'zlardan boshlab nomzod yig'amiz (tezlik uchun)
            posts = sorted((self.tok[t] for t in set(rtoks)
                            if len(t) >= 3 and t in self.tok), key=len)
            for p in posts[:3]:
                if len(p) > 3000 and cnt:
                    break
                for i in p[:3000]:
                    cnt[i] += 1
            sm = SequenceMatcher(None)
            sm.set_seq2(relaxed)
            scored = []
            for i, _ in cnt.most_common(80):
                rel, forms, inn, nom = self.items[i]
                sm.set_seq1(rel)
                if sm.real_quick_ratio() < 0.6:
                    continue
                r = sm.ratio()
                if r > best_hint[0]:
                    best_hint = (r, nom)
                if r >= FUZZY_THRESHOLD and form_state(qforms, forms) != "diff":
                    scored.append((r, i))
            if scored:
                scored.sort(reverse=True)
                r1, i1 = scored[0]
                inn1 = self.items[i1][2]
                other = next((r for r, i in scored[1:] if self.items[i][2] != inn1), None)
                if other is not None and r1 - other < FUZZY_MARGIN:
                    amb = amb or (self.items[i1][3], round(r1, 2), self.cands(
                        [i for r, i in scored if r1 - r < FUZZY_MARGIN]))
                else:
                    q = form_state(qforms, self.items[i1][1])
                    extra = ", tashkiliy shakl ko'rsatilmagan" if q == "unknown" else ""
                    return (inn1, f"taxminiy (imlo farqi{extra})", self.items[i1][3],
                            round(r1, 2), "taxminiy", [])

        if formmis:
            item, score, cl = formmis
            if ALLOW_FORM_MISMATCH:
                return (item[2], "taxminiy (tashkiliy shakl farq qiladi)", item[3],
                        score, "taxminiy", cl)
            return (None, "noaniq (nomi bir xil, tashkiliy shakli farq qiladi)",
                    item[3], score, "noaniq", cl)
        if amb:
            return (None, "noaniq (bir necha mos nomzod/INN)", amb[0], amb[1],
                    "noaniq", amb[2])
        return (None, "topilmadi", best_hint[1], round(best_hint[0], 2),
                "topilmadi", [])


def detect_ref(head):
    """header=None bilan o'qilgan dastlabki qatorlardan INN va NOM ustunlarini topadi."""
    for i in range(len(head)):
        inn = nom = None
        for j, v in enumerate(head.iloc[i].tolist()):
            h = lat(v)
            if not h:
                continue
            if inn is None and re.search(r"\b(inn|stir)\b", h):
                inn = j
            elif nom is None and re.search(r"name|nomi|naimen|korxona|firma", h):
                nom = j
        if inn is not None and nom is not None:
            return i, inn, nom
    return None


def load_ref(path):
    ref = Ref()
    xl = pd.ExcelFile(path)
    total = 0
    want = {fkey(n) for n in REF_SHEETS}
    sheets = [n for n in xl.sheet_names if not want or fkey(n) in want]
    if want and not sheets:
        print(f"  {REF_SHEETS} varag'i topilmadi. Mavjud: {xl.sheet_names}")
        return None
    for name in xl.sheet_names:
        if name not in sheets:
            print(f"  [{name}] o'tkazildi (REF_SHEETS da yo'q)")
            continue
        head = pd.read_excel(xl, sheet_name=name, header=None, nrows=15, dtype=str)
        found = detect_ref(head)
        if not found:
            print(f"  [{name}] INN/NAME ustunlari yo'q - o'tkazildi")
            continue
        hdr, c_inn, c_nom = found
        df = pd.read_excel(xl, sheet_name=name, header=hdr, dtype=str)
        n0 = len(ref.items)
        for inn, nom in zip(df.iloc[:, c_inn].tolist(), df.iloc[:, c_nom].tolist()):
            if nom is None or _isnan(nom):
                continue
            ref.add(nom, clean_inn(inn))
        print(f"  [{name}] ustunlar: INN={df.columns[c_inn]}, NOM={df.columns[c_nom]} "
              f"| qo'shildi: {len(ref.items) - n0}")
        total += len(ref.items) - n0
    return ref if total else None


# ------------------------------------------------- INN yoziladigan fayl
def find_target_columns(ws):
    """(sarlavha_qatori, nom, inn|None, hudud|None, oxirgi_ustun) yoki None."""
    prio = [("korxona", 0), ("naimen", 0), ("name", 1), ("firma", 1),
            ("tashkilot", 1), ("nomi", 2)]
    for r in range(1, 11):
        cells = [c for c in ws[r] if c.value not in (None, "")]
        if len(cells) < 2:
            continue
        nom, inn, hud, last = None, None, None, 0
        for c in cells:
            h = lat(c.value)
            last = max(last, c.column)
            if inn is None and re.search(r"\b(inn|stir)\b", h):
                inn = c.column
            if hud is None and re.search(r"hudud|viloyat|oblast", h):
                hud = c.column
            if re.search(r"tuman|hudud|viloyat|mfy|mahalla", h):
                continue
            for word, p in prio:
                if word in h and (nom is None or p < nom[0]):
                    nom = (p, c.column)
        if nom:
            return r, nom[1], inn, hud, last
    return None


FILLS = {
    "taxminiy": PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC"),
    "noaniq": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
    "topilmadi": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
}


def process_sheet(ws, ref, stats, agg):
    cols = find_target_columns(ws)
    if not cols:
        return None
    hdr, c_nom, c_inn, c_hud, last = cols
    created = c_inn is None
    if created:
        c_inn = last + 1
        h = ws.cell(hdr, c_inn)
        h.value = HEADER_INN
        h._style = copy(ws.cell(hdr, last)._style)
        ws.column_dimensions[get_column_letter(c_inn)].width = 14
        for r in range(hdr + 1, ws.max_row + 1):
            ws.cell(r, c_inn)._style = copy(ws.cell(r, last)._style)
            ws.cell(r, c_inn).alignment = Alignment(horizontal="center",
                                                    vertical="center")
        if ws.auto_filter.ref:
            cr = CellRange(ws.auto_filter.ref)
            if c_inn > cr.max_col:
                ws.auto_filter.ref = CellRange(
                    min_col=cr.min_col, min_row=cr.min_row,
                    max_col=c_inn, max_row=cr.max_row).coord

    hud_key = fkey(HUDUD_FILTER) if HUDUD_FILTER else ""
    st = stats.setdefault(ws.title, Counter())
    for r in range(hdr + 1, ws.max_row + 1):
        name = ws.cell(r, c_nom).value
        if name in (None, "") or not str(name).strip():
            continue
        if ONLY_VISIBLE_ROWS and ws.row_dimensions[r].hidden:
            st["o'tkazildi (filtr yashirgan)"] += 1
            continue
        if hud_key and c_hud and hud_key not in fkey(ws.cell(r, c_hud).value):
            st["o'tkazildi (boshqa hudud)"] += 1
            continue
        cell = ws.cell(r, c_inn)
        if cell.value not in (None, "") and not OVERWRITE_EXISTING:
            st["mavjud (tegilmadi)"] += 1
            continue
        inn, holat, found, score, level, cl = ref.lookup(str(name))
        st[level] += 1
        agg.setdefault((str(name).strip(), holat, found, score, inn or "",
                        "\n".join(cl[:15]), len(cl)), [0])[0] += 1
        if inn:
            cell.value = inn if INN_AS_TEXT else int(inn)
            cell.number_format = "@" if INN_AS_TEXT else "0"
            cell.alignment = Alignment(horizontal="center", vertical="center")
        if HIGHLIGHT:
            if level in FILLS:
                cell.fill = FILLS[level]
            elif (inn and cell.fill.fill_type == "solid"
                  and str(cell.fill.start_color.rgb)[-6:] in ("F8CBAD", "FFF2CC")):
                cell.fill = copy(ws.cell(r, last).fill)
        if len(st) and sum(st.values()) % 20000 == 0:
            print(f"   ... {sum(st.values())} qator")
    return c_inn


def add_report(wb, agg):
    if "INN_hisobot" in wb.sheetnames:
        del wb["INN_hisobot"]
    ws = wb.create_sheet("INN_hisobot")
    ws.append(["Korxona nomi (SSP fayl)", "Holat", "Topilgan/yaqin nom (110 fayl)",
               "O'xshashlik", "Yozilgan INN", "Qatorlar soni",
               "Nomzodlar soni", "Nomzod INN lar (110 fayl) - qo'lda tanlang"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    rows = [[n, h, f, s, i, cnt[0], k or None, c or None]
            for (n, h, f, s, i, c, k), cnt in agg.items()
            if not h.startswith("aniq")]
    rows.sort(key=lambda x: (x[1].split(" ")[0], -x[5]))
    for row in rows:
        ws.append(row)
        ws.cell(ws.max_row, 8).alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([44, 52, 44, 12, 14, 12, 12, 70], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ------------------------------------------------------------- tekshiruv
def style_sig(c):
    return (repr(c.font), repr(c.fill), repr(c.alignment),
            repr(c.border), c.number_format)


def verify(src, dst, skip_cols):
    a, b = load_workbook(src), load_workbook(dst)
    problems = []
    for wa in a.worksheets:
        if wa.title == "INN_hisobot":
            continue
        if wa.title not in b.sheetnames:
            problems.append(f"{wa.title}: varaq yo'q")
            continue
        wb_ = b[wa.title]
        skip = skip_cols.get(wa.title)
        for row in wa.iter_rows():
            for ca in row:
                if ca.column == skip:
                    continue
                cb = wb_.cell(ca.row, ca.column)
                if ca.value != cb.value:
                    problems.append(f"{wa.title}!{ca.coordinate}: qiymat o'zgargan")
                if style_sig(ca) != style_sig(cb):
                    problems.append(f"{wa.title}!{ca.coordinate}: format o'zgargan")
            if len(problems) > 20:
                return problems
        skip_letter = get_column_letter(skip) if skip else None
        for k, d in wa.column_dimensions.items():
            if k == skip_letter:
                continue
            if d.width != wb_.column_dimensions[k].width:
                problems.append(f"{wa.title}: {k} ustun kengligi o'zgargan")
        if wa.merged_cells.ranges != wb_.merged_cells.ranges:
            problems.append(f"{wa.title}: birlashgan kataklar farq qiladi")
    return problems


# ----------------------------------------------------------------- fayllar
def list_excels():
    return sorted(f for f in os.listdir(BASE_DIR)
                  if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"))


def find_file(name, title):
    k = fkey(name)
    for f in list_excels():
        if fkey(f) == k:
            return os.path.join(BASE_DIR, f)
    files = list_excels()
    print(f"'{name}' topilmadi. {title}")
    for i, f in enumerate(files, 1):
        mb = os.path.getsize(os.path.join(BASE_DIR, f)) / 1024 / 1024
        print(f"  {i}. {f}  ({mb:.1f} MB)")
    try:
        return os.path.join(BASE_DIR, files[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def main():
    target = sys.argv[1] if len(sys.argv) > 1 else find_file(
        TARGET_FILE, "INN yoziladigan (SSP) faylni tanlang:")
    ref_path = sys.argv[2] if len(sys.argv) > 2 else find_file(
        REF_FILE, "INN olinadigan (110) faylni tanlang:")
    if not target or not ref_path:
        print("Fayllar tanlanmadi.")
        return
    print("SSP fayl :", os.path.basename(target))
    print("INN manba:", os.path.basename(ref_path))

    print("INN manba fayl o'qilmoqda...")
    ref = load_ref(ref_path)
    if ref is None:
        print("INN manba faylda INN va NOM ustunlari topilmadi.")
        return
    print(f"Manba: {len(ref.items)} korxona")

    print("SSP fayl o'qilmoqda (katta bo'lsa biroz kutiladi)...")
    wb = load_workbook(target)
    stats, agg, skip_cols = {}, {}, {}
    for ws in wb.worksheets:
        if any(k in fkey(ws.title) for k in SKIP_SHEETS):
            continue
        col = process_sheet(ws, ref, stats, agg)
        if col:
            skip_cols[ws.title] = col
            print(f"[{ws.title}] INN ustuni: {get_column_letter(col)} | {dict(stats[ws.title])}")
        else:
            print(f"[{ws.title}] korxona nomi ustuni topilmadi - o'tkazildi")
    if not skip_cols:
        print("Hech bir varaq ishlanmadi.")
        return

    if ADD_REPORT_SHEETS:
        add_report(wb, agg)

    out = target if IN_PLACE else os.path.join(BASE_DIR, OUTPUT_NAME)
    ref_copy = target
    if IN_PLACE:
        ref_copy = (os.path.splitext(target)[0]
                    + f"_zaxira_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(target, ref_copy)
        print("Zaxira nusxa:", os.path.basename(ref_copy))
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"natija_INN_{datetime.now():%H%M%S}.xlsx")
        print("Fayl Excelda ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)
    print("Saqlandi:", out)
    if ADD_REPORT_SHEETS:
        print("  INN_hisobot - taxminiy / noaniq / topilmadi korxonalar (nomlar bo'yicha)")

    if VERIFY:
        try:
            problems = verify(ref_copy, out, skip_cols)
        except Exception as e:
            print("Tekshiruvni bajarib bo'lmadi:", e)
            return
        if problems:
            print("TEKSHIRUV: MUAMMO topildi:")
            for p in problems[:20]:
                print("  -", p)
        else:
            print("TEKSHIRUV: OK - INN ustunidan boshqa hamma ustun va format bir xil.")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
