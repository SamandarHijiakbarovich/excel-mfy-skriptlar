"""
Svet (natija) jadvaliga "MFY STIR" ni Ҳудуд faylining "Бирликлар" varag'idan olib o'tadi.

Qidiruv tartibi (har bir qator uchun):
  1. Tuman ichida: MFY nomi aniq mos           -> aniq
  2. Tuman ichida: KAP nomi aniq mos           -> taxminiy
  3. Tuman ichida: MFY nomi imlo farqi bilan (>= FUZZY_THRESHOLD)   -> taxminiy
  4. Tuman ichida: KAP nomi imlo farqi bilan   -> taxminiy
  5. Tuman topilmasa: butun baza bo'yicha 1-4  -> taxminiy (butun baza)
Bir nomga bir necha STIR mos kelsa, noto'g'ri STIR yozmaslik uchun YOZILMAYDI.

Natija: asl fayl formatlari saqlanadi (faqat "MFY STIR" ustuni o'zgaradi),
oxirida "STIR_hisobot" va "STIR_tuman" varaqlari qo'shiladi.

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
SVET_KEYWORD = "svet"          # natija (Svet) fayl nomida shu bo'lsin
XUDUD_KEYWORD = "hudud"        # Ҳудуд fayl nomida shu bo'lsin
XUDUD_SHEET = "Бирликлар"      # topilmasa, ustunlari mos varaq avtomatik topiladi
OUTPUT_NAME = "natija_MFY_STIR_Mukammal.xlsx"
SVET_FILE = "Svet natija_MFY_STIR_Mukammal.xlsx"   # aniq fayl nomi; bo'sh bo'lsa SVET_KEYWORD bo'yicha qidiriladi
IN_PLACE = True            # True: STIR shu faylning O'ZIGA yoziladi (zaxira nusxa bilan)
SKIP_SHEETS = ["svod", "stirhisobot", "stirtuman"]         # nomida shu bo'lgan varaqlarga tegilmaydi (Свод)

FUZZY_THRESHOLD = 0.80         # tuman ichida o'xshashlik chegarasi
GLOBAL_FUZZY = 0.85            # butun baza bo'yicha (xavfliroq, shuning uchun qattiqroq)
FUZZY_MARGIN = 0.03            # eng yaxshi nomzod ikkinchisidan shuncha ustun bo'lsin
GLOBAL_FALLBACK_ALWAYS = False  # True: tuman topilgan bo'lsa ham, topilmagan MFYni
                                # butun bazadan qidirish (noto'g'ri STIR xavfi bor)
OVERWRITE_EXISTING = False     # False: "MFY STIR" to'ldirilgan kataklarga tegilmaydi
STIR_AS_TEXT = True            # True: matn formatida (bosh nollar/ko'rinish uchun xavfsiz)
HIGHLIGHT = True               # taxminiy = sariq, noaniq/topilmadi = to'q sariq
ADD_REPORT_SHEETS = True
VERIFY = True                  # saqlagach, asl fayl bilan solishtirib tekshirish
CITY_CODE_PREFIX = "17224"     # MXOBT shu bilan boshlansa - shahar (Termiz sh. uchun)
HEADER_STIR = "MFY STIR"
# ======================================================

# ---------------------------------------------------------------- matn
LOOK = str.maketrans("acekmoptxyh", "асекмортхун")
CYR2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
    "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "",
    "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "ғ": "g",
    "қ": "k", "ҳ": "h",
}
APOS = re.compile(r"[\"'`ʻʼ’‘´“”«»]")
STOP_TUMAN = {"tet", "shet", "et", "tuman", "tumani", "shahar", "shahri",
              "sh", "gorod", "rn", "r", "n", "t"}
MFY_NOISE = re.compile(
    r"\bm\.\s*f\.\s*y\.?|\bq\.\s*a\.\s*p\.?|\bsh\.\s*f\.\s*y\.?|"
    r"\b(?:mfy|mfj|mfi|msg|qfy|kfy|qap|shfy|maxallasi|maxalla|mahallasi|"
    r"mahalla|massivi|massiv|fuqarolar|yigini|shaharcha|mchj)\b")


def nfc(s):
    return unicodedata.normalize("NFC", str(s))


def norm(v):
    return " ".join(nfc(v if v is not None and not _isnan(v) else "").casefold().split())


def _isnan(v):
    try:
        return bool(pd.isna(v))
    except Exception:
        return False


def translit(v):
    """Kichik harf, lotinga o'tkazilgan, tutuq belgilari olib tashlangan."""
    s = "".join(CYR2LAT.get(ch, ch) for ch in norm(v))
    return APOS.sub("", s)


def reduce_key(s):
    """Imlo variantlarini tenglashtiradi: x=h, q=k, w=v, qo'sh harf -> bitta."""
    s = re.sub(r"[^a-z0-9]", "", s)
    s = s.replace("x", "h").replace("q", "k").replace("w", "v")
    return re.sub(r"(.)\1+", r"\1", s)


def mfy_key(raw):
    return reduce_key(MFY_NOISE.sub(" ", translit(raw)))


def tuman_key(raw):
    """-> (kalit, ishora). ishora: 'city' | 'tuman' | None"""
    s = translit(raw)
    if re.search(r"\b(?:shet|shahar|shahri|sh|gorod)\b", s):
        hint = "city"
    elif re.search(r"\b(?:tet|tuman|tumani)\b", s):
        hint = "tuman"
    else:
        hint = None
    toks = [t for t in re.findall(r"[a-z0-9]+", s) if t not in STOP_TUMAN]
    return reduce_key("".join(toks)), hint


def fkey(v):
    return re.sub(r"[^a-z0-9]", "", translit(v))


def clean_stir(v):
    if v is None or _isnan(v):
        return ""
    s = re.sub(r"\.0+$", "", str(v).strip())
    return s if s.isdigit() and len(s) >= 5 else ""


def top(cnt):
    return cnt.most_common(1)[0][0]


# ------------------------------------------------------- moslashtirish
def match_exact(pool, mk):
    cnt = pool.get(mk)
    if not cnt:
        return None
    return (None if len(cnt) > 1 else top(cnt)), mk, 1.0


def match_fuzzy(pool, mk, thr):
    if len(mk) < 4:
        return None
    sm = SequenceMatcher(None)
    sm.set_seq2(mk)
    scored = []
    for k in pool:
        sm.set_seq1(k)
        if sm.real_quick_ratio() < thr or sm.quick_ratio() < thr:
            continue
        r = sm.ratio()
        if r >= thr:
            scored.append((r, k))
    if not scored:
        return None
    scored.sort(reverse=True)
    r1, k1 = scored[0]
    if len(pool[k1]) > 1:
        return None, k1, round(r1, 2)
    s1 = top(pool[k1])
    for r2, k2 in scored[1:]:
        if len(pool[k2]) > 1 or top(pool[k2]) != s1:
            if r1 - r2 < FUZZY_MARGIN:
                return None, k1, round(r1, 2)
            break
    return s1, k1, round(r1, 2)


def nearest(pool, mk):
    best = (0.0, "")
    sm = SequenceMatcher(None)
    sm.set_seq2(mk)
    for k in pool:
        sm.set_seq1(k)
        if sm.real_quick_ratio() <= best[0]:
            continue
        r = sm.ratio()
        if r > best[0]:
            best = (r, k)
    return best


class Base:
    def __init__(self):
        self.mfy, self.kap = {}, {}           # gid -> {kalit: Counter(STIR)}
        self.gmfy, self.gkap = {}, {}         # butun baza
        self.groups = {}                      # tuman kaliti -> {gid}
        self.disp_mfy, self.disp_kap = {}, {}
        self.cache, self._pools = {}, {}
        self.rows = 0

    def add(self, tuman, mxobt, mfy, kap, stir):
        mk, kk = mfy_key(mfy), mfy_key(kap)
        tk, hint = tuman_key(tuman)
        if not stir or not tk or not (mk or kk):
            return
        mx = "" if _isnan(mxobt) else str(mxobt).strip()
        gid = (tk, hint == "city" or mx.startswith(CITY_CODE_PREFIX))
        self.groups.setdefault(tk, set()).add(gid)
        if mk:
            self.mfy.setdefault(gid, {}).setdefault(mk, Counter())[stir] += 1
            self.gmfy.setdefault(mk, Counter())[stir] += 1
            self.disp_mfy.setdefault(mk, str(mfy).strip())
        if kk:
            self.kap.setdefault(gid, {}).setdefault(kk, Counter())[stir] += 1
            self.gkap.setdefault(kk, Counter())[stir] += 1
            self.disp_kap.setdefault(kk, str(kap).strip())
        self.rows += 1

    def pool(self, store, gids):
        key = (id(store), tuple(sorted(gids)))
        if key not in self._pools:
            if len(gids) == 1:
                self._pools[key] = store.get(gids[0], {})
            else:
                out = {}
                for g in gids:
                    for k, c in store.get(g, {}).items():
                        out.setdefault(k, Counter()).update(c)
                self._pools[key] = out
        return self._pools[key]

    def _resolve(self, raw):
        """-> (asosiy guruhlar, qo'shni guruhlar, tuman kaliti)"""
        tk, hint = tuman_key(raw)
        if not tk:
            return None, [], tk
        cands = self.groups.get(tk)
        if not cands:
            best = max(((SequenceMatcher(None, tk, k).ratio(), k)
                        for k in self.groups), default=(0, None))
            if best[0] < 0.8:
                return None, [], tk
            cands = self.groups[best[1]]
        cands = sorted(cands)
        if hint == "city":
            pick = [g for g in cands if g[1]]
        elif hint == "tuman":
            pick = [g for g in cands if not g[1]]
        else:
            pick = []
        main = pick or cands
        return main, [g for g in cands if g not in main], tk

    def resolve_groups(self, raw):
        main, _, tk = self._resolve(raw)
        return main, tk

    def lookup(self, tuman_raw, mfy_raw):
        """-> (stir|None, holat, topilgan_nom, o'xshashlik, daraja)"""
        ck = (tuman_raw, mfy_raw)
        if ck not in self.cache:
            self.cache[ck] = self._lookup(tuman_raw, mfy_raw)
        return self.cache[ck]

    def _lookup(self, tuman_raw, mfy_raw):
        mk = mfy_key(mfy_raw)
        if not mk:
            return None, "MFY nomi bo'sh", "", 0.0, "bosh"
        gids, rest, _ = self._resolve(tuman_raw)
        scopes = []
        if gids:
            scopes.append(("tuman", self.pool(self.mfy, gids),
                           self.pool(self.kap, gids), FUZZY_THRESHOLD))
        if gids and rest:   # masalan: "Термиз ТЭТ" qatori, MFY esa Termiz shahri guruhida
            scopes.append(("qoshni", self.pool(self.mfy, rest),
                           self.pool(self.kap, rest), FUZZY_THRESHOLD))
        if not gids or GLOBAL_FALLBACK_ALWAYS:
            scopes.append(("butun baza", self.gmfy, self.gkap, GLOBAL_FUZZY))

        amb = None
        for scope, mp, kp, thr in scopes:
            tries = [("exact", "MFY", mp, self.disp_mfy),
                     ("exact", "KAP", kp, self.disp_kap),
                     ("fuzzy", "MFY", mp, self.disp_mfy),
                     ("fuzzy", "KAP", kp, self.disp_kap)]
            for mode, label, pl, disp in tries:
                r = match_exact(pl, mk) if mode == "exact" else match_fuzzy(pl, mk, thr)
                if r is None:
                    continue
                stir, key, score = r
                name = disp.get(key, key)
                if stir is None:
                    amb = amb or (name, score)
                    continue
                if scope == "tuman" and label == "MFY" and mode == "exact":
                    return stir, "aniq (tuman + MFY nomi)", name, 1.0, "aniq"
                what = "MFY nomi" if label == "MFY" else "KAP nomi"
                how = "imlo farqi bilan" if mode == "fuzzy" else "aniq"
                where = {"tuman": "tuman ichida",
                         "qoshni": "shu nomli boshqa guruhda (shahar/tuman)"
                         }.get(scope, "butun baza, tuman topilmadi")
                return (stir, f"taxminiy ({what}, {how}, {where})", name,
                        score, "taxminiy")
        if amb:
            return (None, "noaniq (bir necha mos nomzod/STIR)", amb[0], amb[1],
                    "noaniq")
        near = nearest(self.pool(self.mfy, gids) if gids else self.gmfy, mk)
        disp = self.disp_mfy.get(near[1], near[1])
        why = "topilmadi" if gids else "topilmadi (tuman ham topilmadi)"
        return None, why, disp, round(near[0], 2), "topilmadi"


# --------------------------------------------------------- Ҳудуд fayl
def detect_xudud(head):
    """head: header=None bilan o'qilgan dastlabki qatorlar."""
    for i in range(len(head)):
        cols = {}
        for j, v in enumerate(head.iloc[i].tolist()):
            raw = norm(v)
            if not raw:
                continue
            cy = raw.translate(LOOK)
            is_mfy = "mfy" in raw or "мфй" in cy or "мфи" in cy
            has_nom = "nom" in raw or "ном" in cy
            if "stir" in raw or "стир" in cy:
                cols.setdefault("stir", j)
            elif is_mfy and has_nom:
                cols.setdefault("mfy", j)
            elif "tuman" in raw or "туман" in cy:
                cols.setdefault("tuman", j)
            elif ("kap" in raw or "кап" in cy) and has_nom:
                cols.setdefault("kap", j)
            elif "мхобт" in cy or "mxobt" in raw or "mhobt" in raw:
                cols.setdefault("mxobt", j)
        if {"stir", "mfy", "tuman"} <= cols.keys():
            return i, cols
    return None


def load_xudud(path):
    xl = pd.ExcelFile(path)
    want = fkey(XUDUD_SHEET)
    order = sorted(xl.sheet_names, key=lambda n: 0 if fkey(n) == want else 1)
    for name in order:
        head = pd.read_excel(xl, sheet_name=name, header=None, nrows=15, dtype=str)
        found = detect_xudud(head)
        if not found:
            continue
        hdr, cols = found
        df = pd.read_excel(xl, sheet_name=name, header=hdr, dtype=str)
        print(f"Ҳудуд varag'i: [{name}] | sarlavha qatori: {hdr + 1} | ustunlar: "
              + ", ".join(f"{k}={df.columns[v]}" for k, v in cols.items()))
        return df, cols
    return None, None


def build_base(df, cols):
    get = lambda k: (df.iloc[:, cols[k]].tolist() if k in cols
                     else [None] * len(df))
    base = Base()
    for t, mx, m, k, s in zip(get("tuman"), get("mxobt"), get("mfy"),
                              get("kap"), get("stir")):
        base.add(t, mx, m, k, clean_stir(s))
    return base


# ----------------------------------------------------- Svet (natija) fayl
def find_svet_columns(ws):
    for r in range(1, 11):
        tum = mfy = stir = None
        last = 0
        for c in ws[r]:
            if c.value in (None, ""):
                continue
            last = max(last, c.column)
            raw = norm(c.value)
            cy = raw.translate(LOOK)
            is_mfy = "mfy" in raw or "мфй" in cy or "мфи" in cy
            is_stir = "stir" in raw or "стир" in cy
            if tum is None and ("tuman" in raw or "туман" in cy):
                tum = c.column
            elif is_mfy and is_stir and stir is None:
                stir = c.column
            elif (is_mfy and mfy is None and (
                    "nom" in raw or "ном" in cy
                    or raw.replace(" ", "") in ("мфй", "мфи", "mfy")
                    or cy.replace(" ", "") in ("мфй", "мфи"))):
                mfy = c.column
        if tum and mfy:
            return r, tum, mfy, stir, last
    return None


FILLS = {
    "taxminiy": PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC"),
    "noaniq": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
    "topilmadi": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
}


def process_sheet(ws, base, stats, agg):
    cols = find_svet_columns(ws)
    if not cols:
        return None
    hdr, c_tum, c_mfy, c_stir, last = cols
    created = c_stir is None
    if created:
        c_stir = last + 1
        h = ws.cell(hdr, c_stir)
        h.value = HEADER_STIR
        h._style = copy(ws.cell(hdr, c_mfy)._style)
        ws.column_dimensions[get_column_letter(c_stir)].width = 16
        for r in range(hdr + 1, ws.max_row + 1):
            ws.cell(r, c_stir)._style = copy(ws.cell(r, c_mfy)._style)
            ws.cell(r, c_stir).alignment = Alignment(
                horizontal="center", vertical="center")
        if ws.auto_filter.ref:
            cr = CellRange(ws.auto_filter.ref)
            if c_stir > cr.max_col:
                ws.auto_filter.ref = CellRange(
                    min_col=cr.min_col, min_row=cr.min_row,
                    max_col=c_stir, max_row=cr.max_row).coord

    st = stats.setdefault(ws.title, Counter())
    for r in range(hdr + 1, ws.max_row + 1):
        mfy_raw = ws.cell(r, c_mfy).value
        if mfy_raw in (None, "") or not str(mfy_raw).strip():
            continue
        cell = ws.cell(r, c_stir)
        tuman_raw = ws.cell(r, c_tum).value
        if cell.value not in (None, "") and not OVERWRITE_EXISTING:
            st["mavjud (tegilmadi)"] += 1
            # hisobotda oldingi taxminiy qatorlar yo'qolmasin
            _, h0, n0, s0, lv0 = base.lookup(tuman_raw, mfy_raw)
            if lv0 != "aniq":
                agg.setdefault((str(tuman_raw), str(mfy_raw),
                                h0 + " [oldingi natija saqlandi]", n0, s0,
                                str(cell.value)), [0])[0] += 1
            continue
        stir, holat, name, score, level = base.lookup(tuman_raw, mfy_raw)
        st[level] += 1
        agg.setdefault((str(tuman_raw), str(mfy_raw), holat, name, score,
                        stir or ""), [0])[0] += 1
        if stir:
            cell.value = stir if STIR_AS_TEXT else int(stir)
            cell.number_format = "@" if STIR_AS_TEXT else "0"
            cell.alignment = Alignment(horizontal="center", vertical="center")
        if HIGHLIGHT:
            if level in FILLS:
                cell.fill = FILLS[level]
            elif (stir and c_stir > 1 and cell.fill.fill_type == "solid"
                  and str(cell.fill.start_color.rgb)[-6:] in ("F8CBAD", "FFF2CC")):
                cell.fill = copy(ws.cell(r, c_mfy).fill)
    return c_stir


def add_report_sheets(wb, agg, base, svet_tumans):
    for name in ("STIR_hisobot", "STIR_tuman"):
        if name in wb.sheetnames:
            del wb[name]
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")

    def sheet(title, header, rows, widths):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill = head_font, head_fill
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
        for row in rows:
            ws.append(row)
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    rows = [[t, m, h, n, s, stir, cnt[0]]
            for (t, m, h, n, s, stir), cnt in agg.items()
            if not h.startswith("aniq")]
    rows.sort(key=lambda x: (-x[6], x[2]))
    sheet("STIR_hisobot",
          ["Tuman (natija)", "MFY nomi (natija)", "Holat", "Topilgan/yaqin nom",
           "O'xshashlik", "Yozilgan STIR", "Qatorlar soni"],
          rows, [22, 34, 48, 34, 12, 14, 12])

    trows = []
    for raw in sorted(svet_tumans):
        gids, tk = base.resolve_groups(raw)
        if gids:
            size = sum(len(base.mfy.get(g, {})) for g in gids)
            trows.append([raw, tk, ", ".join(f"{g[0]}{'/shahar' if g[1] else ''}"
                                             for g in gids), size])
        else:
            trows.append([raw, tk, "TOPILMADI", 0])
    sheet("STIR_tuman", ["Tuman (natija)", "Kalit", "Ҳудуддagi mos tuman",
                         "Shu tumandagi MFYlar soni"], trows, [24, 18, 34, 18])


# ------------------------------------------------------------- tekshiruv
def style_sig(c):
    return (repr(c.font), repr(c.fill), repr(c.alignment),
            repr(c.border), c.number_format)


def verify(src, dst, skip_cols):
    a, b = load_workbook(src), load_workbook(dst)
    problems = []
    for wa in a.worksheets:
        if wa.title in ("STIR_hisobot", "STIR_tuman"):
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
            if k == skip_letter:      # MFY STIR ustuni kengligini o'zimiz belgilaymiz
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


def find_exact(name):
    k = fkey(name)
    for f in list_excels():
        if fkey(f) == k:
            return os.path.join(BASE_DIR, f)
    print(f"'{name}' papkada topilmadi. Mavjud Excel fayllar:")
    for f in list_excels():
        print("  ", f)
    return None


def pick(keyword, title, exclude=("natija",)):
    files = [f for f in list_excels()
             if keyword in fkey(f) and not any(e in fkey(f) for e in exclude)]
    if len(files) == 1:
        return os.path.join(BASE_DIR, files[0])
    cand = files or list_excels()
    print(title)
    for i, f in enumerate(cand, 1):
        mb = os.path.getsize(os.path.join(BASE_DIR, f)) / 1024 / 1024
        print(f"  {i}. {f}  ({mb:.1f} MB)")
    try:
        return os.path.join(BASE_DIR, cand[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


# -------------------------------------------------------------------- main
def main():
    svet = sys.argv[1] if len(sys.argv) > 1 else (
        find_exact(SVET_FILE) if SVET_FILE else pick(
            SVET_KEYWORD, "Natija (Svet) faylni tanlang:"))
    xudud = sys.argv[2] if len(sys.argv) > 2 else pick(
        XUDUD_KEYWORD, "Ҳудуд faylni tanlang:")
    if not svet or not xudud:
        print("Fayllar tanlanmadi.")
        return
    print("Svet :", os.path.basename(svet))
    print("Ҳудуд:", os.path.basename(xudud))

    print("Ҳудуд fayl o'qilmoqda...")
    df, cols = load_xudud(xudud)
    if df is None:
        print("Ҳудуд faylda STIR + MFY nomi + Tuman ustunlari topilmadi.")
        return
    base = build_base(df, cols)
    print(f"Baza: {base.rows} qator, {len(base.groups)} tuman, "
          f"{len(base.gmfy)} MFY nomi, {len(base.gkap)} KAP nomi")

    print("Svet fayl o'qilmoqda...")
    wb = load_workbook(svet)
    stats, agg, skip_cols, tumans = {}, {}, {}, set()
    for ws in wb.worksheets:
        if any(k in fkey(ws.title) for k in SKIP_SHEETS):
            print(f"[{ws.title}] o'tkazib yuborildi (SKIP_SHEETS)")
            continue
        c = find_svet_columns(ws)
        if c:
            for r in range(c[0] + 1, ws.max_row + 1):
                v = ws.cell(r, c[1]).value
                if v not in (None, ""):
                    tumans.add(str(v))
        col = process_sheet(ws, base, stats, agg)
        if col:
            skip_cols[ws.title] = col
            print(f"[{ws.title}] {dict(stats[ws.title])}")
        else:
            print(f"[{ws.title}] 'tuman' va 'MFY nomi' ustunlari yo'q - o'tkazildi")
    if not skip_cols:
        print("Hech bir varaq ishlanmadi.")
        return

    bad = [t for t in sorted(tumans) if not base.resolve_groups(t)[0]]
    if bad:
        print("DIQQAT: Ҳудуддan topilmagan tumanlar (butun bazadan qidirildi):", bad)

    if ADD_REPORT_SHEETS:
        add_report_sheets(wb, agg, base, tumans)

    out = svet if IN_PLACE else os.path.join(BASE_DIR, OUTPUT_NAME)
    ref = svet
    if IN_PLACE:
        ref = (os.path.splitext(svet)[0]
               + f"_zaxira_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(svet, ref)
        print("Zaxira nusxa:", os.path.basename(ref))
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"natija_STIR_{datetime.now():%H%M%S}.xlsx")
        print("Fayl Excelda ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)
    print("Saqlandi:", out)
    if ADD_REPORT_SHEETS:
        print("  STIR_hisobot - taxminiy/noaniq/topilmadi qatorlar (nomlar bo'yicha)")
        print("  STIR_tuman   - tumanlar Ҳудуд bilan qanday moslangani")

    if VERIFY:
        try:
            problems = verify(ref, out, skip_cols)
        except Exception as e:
            print("Tekshiruvni bajarib bo'lmadi:", e)
            return
        if problems:
            print("TEKSHIRUV: MUAMMO topildi:")
            for p in problems[:20]:
                print("  -", p)
        else:
            print("TEKSHIRUV: OK - 'MFY STIR' dan boshqa hamma ustun va format bir xil.")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
