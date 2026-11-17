"""
1722 faylidan (MFY nomi + tuman bo'yicha) MFY STIR ni topib,
natija_MFY.xlsx dagi "Мулк" va "техника ва бошқалар" varaqlarining
"МФЙ СТИР" ustuniga yozadi.

* Kirill va lotin yozuvi tenglashtiriladi (ШАРҚГУЛИ = SHARQGULI).
* Tuman manzildan olinadi (Ангор тумани -> angor), shuning uchun boshqa
  tumandagi bir xil nomli mahallalar aralashmaydi.
* Imlo farqlari (TALLOSHQON / TALLASHQON) hisobga olinadi, lekin
  ishonchsiz moslik yozilmaydi - hisobot faylida ko'rsatiladi.
* Faqat "МФЙ СТИР" ustuni o'zgaradi. Natija: natija_MFY_STIR.xlsx
* Saqlagandan so'ng skript asl bilan solishtirib tekshiradi.

Oldin main.py ishlagan bo'lishi kerak (natija_MFY.xlsx hosil bo'lgan).
"""
import csv
import difflib
import json
import os
import re
import sys
import traceback
import unicodedata
from datetime import datetime

from openpyxl import load_workbook

# ===================== SOZLAMALAR =====================
SOURCE_KEYWORD = "1722"                  # manba fayl nomida shu bo'lsin
DEST_NAME = "natija_MFY.xlsx"            # main.py natijasi
OUTPUT_NAME = "natija_MFY_STIR.xlsx"
REPORT_NAME = "stir_hisobot.csv"
CACHE_NAME = "1722_indeks.json"          # 63 MB faylni qayta o'qimaslik uchun
SHEET_KEYWORDS = ["мулк", "техника"]
SKIP_SHEETS = ["свод"]
FUZZY_THRESHOLD = 0.84                   # imlo farqi uchun (0..1)
FUZZY_MARGIN = 0.04                      # ikkinchi nomzoddan qanchaga ustun
# ======================================================

BASE = os.path.dirname(os.path.abspath(__file__))
LOOK = str.maketrans("acekmoptxyh", "асекмортхун")

CYR2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
    "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "ъ": "", "ы": "i",
    "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "ғ": "g", "қ": "k",
    "ҳ": "h",
}
NOISE = re.compile(
    r"\b(?:mfy|mfj|mfi|msg|shfy|shfj|qfy|kfy|mahallasi|mahalla|mahallasi|"
    r"fuqarolar|yigini|mchj|tumani|tuman|shahri|shahar)\b")


def nfc(s):
    return unicodedata.normalize("NFC", str(s))


def norm(v):
    return " ".join(nfc(v if v is not None else "").casefold().split())


def cyr(v):
    return norm(v).translate(LOOK)


def name_key(v):
    """Mahalla/tuman nomining yozuvdan mustaqil kaliti."""
    s = norm(v)
    if not s:
        return ""
    s = "".join(CYR2LAT.get(ch, ch) for ch in s)
    s = re.sub(r"[\"'`ʻʼ’‘´“”«»]", "", s)
    s = NOISE.sub(" ", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    s = s.replace("x", "h").replace("q", "k").replace("w", "v")
    return re.sub(r"(.)\1+", r"\1", s)


def ratio(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()


def clean_stir(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return s if s.isdigit() else ""


# ------------------------------------------------------------- fayllar
def list_excels():
    return sorted(
        f for f in os.listdir(BASE)
        if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$"))


def choose(files, title):
    print(title)
    for i, f in enumerate(files, 1):
        mb = os.path.getsize(os.path.join(BASE, f)) / 1024 / 1024
        print(f"  {i}. {f}  ({mb:.1f} MB)")
    try:
        return os.path.join(BASE, files[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def pick_source():
    files = [f for f in list_excels() if norm(SOURCE_KEYWORD) in norm(f)]
    if len(files) == 1:
        return os.path.join(BASE, files[0])
    return choose(files or list_excels(),
                  "Manba (1722) faylni tanlang:")


def pick_dest():
    p = os.path.join(BASE, DEST_NAME)
    if os.path.exists(p):
        return p
    files = [f for f in list_excels()
             if "natija" in norm(f) and "stir" not in norm(f)]
    return choose(files or list_excels(), f"{DEST_NAME} topilmadi. Tanlang:")


# ------------------------------------------------- manba indeksi (1722)
def detect_src(row):
    stir = name = tuman = None
    for i, v in enumerate(row):
        raw = norm(v)
        if not raw:
            continue
        cy = raw.translate(LOOK)
        is_mfy = "mfy" in raw or "мфй" in cy or "мфи" in cy
        if stir is None and is_mfy and ("stir" in raw or "стир" in cy):
            stir = i
        elif name is None and is_mfy and ("nom" in raw or "ном" in cy):
            name = i
        elif tuman is None and ("tuman" in raw or "туман" in cy):
            tuman = i
    return (stir, name, tuman) if stir is not None and name is not None else None


def build_index(path):
    st = os.stat(path)
    sig = [st.st_size, int(st.st_mtime)]
    cache = os.path.join(BASE, CACHE_NAME)
    if os.path.exists(cache):
        try:
            with open(cache, encoding="utf-8") as f:
                d = json.load(f)
            if d.get("sig") == sig:
                print("Indeks keshdan olindi (1722 fayl qayta o'qilmadi).")
                return d["index"]
        except Exception:
            pass

    print("1722 fayl o'qilmoqda - katta fayl, bir necha daqiqa ketishi mumkin...")
    wb = load_workbook(path, read_only=True, data_only=True)
    index = {}
    for ws in wb.worksheets:
        try:
            ws.reset_dimensions()
        except Exception:
            pass
        cols, used = None, 0
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if cols is None:
                if i >= 10:
                    break
                cols = detect_src(row)
                continue
            s_i, n_i, t_i = cols
            get = lambda k: row[k] if k is not None and k < len(row) else None
            nk, sv = name_key(get(n_i)), clean_stir(get(s_i))
            if not nk or not sv:
                continue
            tk = name_key(get(t_i)) or "*"
            d = index.setdefault(tk, {}).setdefault(nk, {})
            d[sv] = d.get(sv, 0) + 1
            used += 1
        if cols is None:
            print(f"  [{ws.title}] MFY STIR/MFY nomi ustunlari yo'q - o'tkazildi")
        else:
            print(f"  [{ws.title}] qatorlar: {used}")
    wb.close()
    with open(cache, "w", encoding="utf-8") as f:
        json.dump({"sig": sig, "index": index}, f, ensure_ascii=False)
    return index


class Matcher:
    def __init__(self, index):
        self.index = index
        self.tumans = [t for t in index if t != "*"]
        self.cache = {}
        self._global = None

    def global_pool(self):
        if self._global is None:
            pool = {}
            for names in self.index.values():
                for nk, cnt in names.items():
                    d = pool.setdefault(nk, {})
                    for s, c in cnt.items():
                        d[s] = d.get(s, 0) + c
            self._global = pool
        return self._global

    def find_tuman(self, tk):
        if not tk:
            return None
        if tk in self.index:
            return tk
        best = max(((ratio(tk, t), t) for t in self.tumans), default=(0, None))
        return best[1] if best[0] >= 0.8 else None

    @staticmethod
    def top(cnt):
        return max(cnt.items(), key=lambda kv: kv[1])[0]

    def lookup(self, tk, nk):
        """-> (stir yoki None, holat, moslik_nomi, oxshashlik)"""
        key = (tk, nk)
        if key not in self.cache:
            self.cache[key] = self._lookup(tk, nk)
        return self.cache[key]

    def _lookup(self, tk, nk):
        rt = self.find_tuman(tk)
        pool = self.index[rt] if rt else self.global_pool()
        by_tuman = rt is not None

        if nk in pool:
            cnt = pool[nk]
            if not by_tuman and len(cnt) > 1:
                return None, "noaniq (tuman yo'q, bir necha STIR)", nk, 1.0
            return self.top(cnt), "aniq", nk, 1.0

        scored = sorted(((ratio(nk, k), k) for k in pool), reverse=True)
        if not scored or scored[0][0] < FUZZY_THRESHOLD:
            best = scored[0] if scored else (0, "")
            return None, "topilmadi", best[1], round(best[0], 2)
        r1, k1 = scored[0]
        s1 = self.top(pool[k1])
        for r2, k2 in scored[1:]:
            if self.top(pool[k2]) != s1:
                if r1 - r2 < FUZZY_MARGIN:
                    return None, "noaniq (2 ta o'xshash mahalla)", k1, round(r1, 2)
                break
        if not by_tuman and len(pool[k1]) > 1:
            return None, "noaniq (tuman yo'q, bir necha STIR)", k1, round(r1, 2)
        return s1, "taxminiy", k1, round(r1, 2)


# ------------------------------------------------- manzil / ustunlar (natija)
TUMAN_RE = re.compile(r"([^\s,]+)\s+(?:тумани|tumani)", re.I)


def tuman_from_address(addr):
    if addr is None:
        return ""
    m = TUMAN_RE.search(nfc(addr))
    return name_key(m.group(1)) if m else ""


def wanted_sheet(title):
    c = cyr(title)
    if any(k in c for k in SKIP_SHEETS):
        return False
    return any(k in c for k in SHEET_KEYWORDS)


def find_dest_columns(ws):
    """(sarlavha_qatori, stir, nom, manzil|None, tuman|None) yoki None."""
    for row in ws.iter_rows(min_row=1, max_row=8):
        stir = nom = addr = tum = None
        for c in row:
            raw = norm(c.value)
            if not raw:
                continue
            cy = raw.translate(LOOK)
            is_mfy = "mfy" in raw or "мфй" in cy or "мфи" in cy
            prop = any(w in cy for w in ("жойлаш", "кўчмас", "кучмас"))
            live = "яшаш" in cy or "yashash" in raw
            if stir is None and is_mfy and ("стир" in cy or "stir" in raw) \
                    and not prop:
                stir = c.column
            elif nom is None and (cy.replace(" ", "") in ("мфйноми", "мфиноми")
                                  or raw.replace(" ", "") == "mfynomi"):
                nom = c.column
            elif addr is None and live and ("манзил" in cy or "manzil" in raw):
                addr = c.column
            elif tum is None and ("туман" in cy or "tuman" in raw) and not prop:
                tum = c.column
        if stir and nom:
            return row[0].row, stir, nom, addr, tum
    return None


# ------------------------------------------------------------- tekshiruv
def style_sig(c):
    return (repr(c.font), repr(c.fill), repr(c.alignment),
            repr(c.border), c.number_format)


def verify(src, dst, target_cols):
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
            if len(problems) > 20:
                return problems
        for k, d in wa.column_dimensions.items():
            if d.width != wb_.column_dimensions[k].width:
                problems.append(f"{wa.title}: {k} ustun kengligi o'zgargan")
        if wa.merged_cells.ranges != wb_.merged_cells.ranges:
            problems.append(f"{wa.title}: birlashgan kataklar farq qiladi")
        if wa.auto_filter.ref != wb_.auto_filter.ref:
            problems.append(f"{wa.title}: filtr o'zgargan")
    return problems


# ------------------------------------------------------------------ asosiy
def main():
    source = pick_source()
    dest = pick_dest()
    if not source or not dest:
        print("Fayllar tanlanmadi.")
        return
    print("Manba:", os.path.basename(source))
    print("Natija:", os.path.basename(dest))

    index = build_index(source)
    if not index:
        print("Manba faylda MFY STIR / MFY nomi ustunlari topilmadi.")
        return
    total = sum(len(v) for v in index.values())
    print(f"Indeks: {len(index)} tuman, {total} mahalla nomi.")
    matcher = Matcher(index)

    print("Natija fayl o'qilmoqda...")
    wb = load_workbook(dest)
    target_cols, report = {}, []
    stats = {}

    for ws in wb.worksheets:
        if not wanted_sheet(ws.title):
            continue
        cols = find_dest_columns(ws)
        if not cols:
            print(f"[{ws.title}] 'МФЙ СТИР' / 'МФЙ номи' ustunlari topilmadi.")
            continue
        hdr, c_stir, c_nom, c_addr, c_tum = cols
        target_cols[ws.title] = c_stir
        st = stats.setdefault(ws.title, {})

        for r in range(hdr + 1, ws.max_row + 1):
            raw_name = ws.cell(r, c_nom).value
            if raw_name in (None, ""):
                continue
            nk = name_key(raw_name)
            if not nk:
                continue
            tk = tuman_from_address(ws.cell(r, c_addr).value) if c_addr else ""
            if not tk and c_tum:
                tk = name_key(ws.cell(r, c_tum).value)

            stir, holat, matched, sc = matcher.lookup(tk, nk)
            key = holat.split(" ")[0]
            st[key] = st.get(key, 0) + 1
            if stir:
                ws.cell(r, c_stir).value = int(stir) if stir.isdigit() else stir
            if key != "aniq":
                report.append([ws.title, r, holat, raw_name, tk, matched, sc,
                               stir or ""])
        print(f"[{ws.title}] {st}")

    if not target_cols:
        print("Hech bir varaq ishlanmadi, fayl saqlanmadi.")
        return

    out = os.path.join(BASE, OUTPUT_NAME)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE, f"natija_MFY_STIR_{datetime.now():%H%M%S}.xlsx")
        print(f"{OUTPUT_NAME} Excelda ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)
    print("Saqlandi:", out)

    if report:
        rep = os.path.join(BASE, REPORT_NAME)
        try:
            with open(rep, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                w.writerow(["varaq", "qator", "holat", "natijadagi_nom",
                            "tuman_kaliti", "topilgan_nom_kaliti",
                            "oxshashlik", "yozilgan_STIR"])
                w.writerows(report)
            print(f"Hisobot: {len(report)} qator -> {rep}")
            print("  taxminiy = imlo farqi bilan topilgan va yozilgan (tekshiring)")
            print("  noaniq/topilmadi = STIR YOZILMADI")
        except PermissionError:
            print("Hisobot fayli ochiq, yozilmadi.")

    try:
        problems = verify(dest, out, target_cols)
    except Exception as e:
        print("Tekshiruvni bajarib bo'lmadi:", e)
        return
    if problems:
        print("TEKSHIRUV: MUAMMO topildi:")
        for p in problems[:20]:
            print("  -", p)
    else:
        print("TEKSHIRUV: OK - 'МФЙ СТИР' dan boshqa hamma narsa bir xil.")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
