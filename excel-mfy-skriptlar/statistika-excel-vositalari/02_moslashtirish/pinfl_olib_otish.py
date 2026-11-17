"""
Ism-familiya (F.I.O) bo'yicha solishtirib, PINFL (JSHSHIR) ni manba fayldan
natija faylga olib o'tadi.

  MANBA (2200 ...)   : sarlavhali, PINFL + IFO(ism) + ADDRESS ustunlari bor.
  NATIJA (Hashlangan): sarlavhasiz, ism va manzil ustunlari bor; shu faylga
                       PINFL yoziladi. HASH ustuniga tegilmaydi.

Solishtirish:
  * kirill/lotin, O'G'LI/O‘G‘LI/OGLI, QIZI/KIZI, x=h, q=k, qo'sh harf tenglashtiriladi;
  * bir xil ismli bir necha odam bo'lsa - MANZIL o'xshashligi bo'yicha ajratiladi;
  * ishonchsiz bo'lsa PINFL yozilmaydi, "Tekshirish" varag'ida nomzodlar ko'rsatiladi.

Natija (yangi fayl, asl fayllar saqlanadi):
  * natija faylning nusxasi + "PINFL" va "Moslik" ustunlari;
  * "Tekshirish" varag'i - taxminiy/noaniq/topilmagan qatorlar.

O'rnatish:  pip install openpyxl
"""
import os
import re
import sys
import time
import traceback
import unicodedata
from collections import Counter, defaultdict
from copy import copy
from datetime import datetime
from difflib import SequenceMatcher

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ===================== SOZLAMALAR =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_FILE = "2200_(1-8).xlsx"        # PINFL olinadigan fayl (sarlavhali)
TARGET_FILE = "Hashlangan.xlsx"        # PINFL yoziladigan fayl (hashlangan)
OUTPUT_NAME = "Hashlangan_PINFL.xlsx"
# Natija faylda ustunlar (sarlavhasiz fayl uchun harf bilan majburlash mumkin):
TARGET_NAME_COL = None     # masalan "B"; None -> avtomatik aniqlash
TARGET_ADDR_COL = None     # masalan "D"; None -> avtomatik aniqlash
FUZZY_THRESHOLD = 0.92     # imlo farqi bilan moslash chegarasi
FUZZY_MARGIN = 0.03
USE_ADDRESS = True         # bir xil ismlilarni manzil bo'yicha ajratish
HIGHLIGHT = True
# ======================================================

HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
CYR2LAT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo",
    "ж": "j", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "x", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": "",
    "ы": "i", "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o", "ғ": "g",
    "қ": "k", "ҳ": "h",
}
APOS = re.compile(r"['`ʻʼ’‘´]")
PATRON = {"ogli", "ugli", "oglu", "ugly", "kizi", "qizi", "kiza", "qiza"}
PATRON_SUF = re.compile(r"(ovich|evich|yevich|ivich|ovna|evna|yevna|ichna|ich|vna)$")
ADDR_STOP = {"surxondaryo", "surhondaryo", "viloyati", "viloyat", "oblast",
             "область", "сурхандарьинская", "tumani", "tuman", "rayon", "район",
             "shahri", "shahar", "shah", "g", "mahallasi", "mahalla", "maxallasi",
             "mfy", "msg", "qfy", "kfy", "siu", "mch", "mchj", "kuchasi", "kochasi",
             "kucha", "ko", "chasi", "uy", "uyi", "dom", "d", "kv", "kvartira",
             "bulim", "bn", "kishlogi", "qishlogi", "qishloq", "kishlok",
             "massivi", "massiv", "махалласи", "кучаси", "дом", "кв", "мфй",
             "район", "кишлоги", "shimoliy", "janubiy"}
HINTS = {
    "pinfl": ["pinfl", "pnfl", "jshshir", "jshr", "jshir", "жшшир", "пинфл"],
    "name": ["ifo", "fio", "fish", "fish", "f.i.sh", "f.i.o", "name", "nomi",
             "фио", "фиш", "talaba"],
    "address": ["address", "manzil", "adres", "манзил", "адрес", "yashash"],
}


def nfc(v):
    return "" if v is None or (isinstance(v, float) and v != v) else \
        unicodedata.normalize("NFC", str(v))


def lat(v):
    s = " ".join(nfc(v).casefold().split())
    return APOS.sub("", "".join(CYR2LAT.get(ch, ch) for ch in s))


def red(t):
    t = re.sub(r"[^a-z0-9]", "", t)
    t = t.replace("x", "h").replace("q", "k").replace("w", "v")
    t = t.replace("ye", "e").replace("yo", "o").replace("yu", "u")
    return re.sub(r"(.)\1+", r"\1", t)


def name_keys(raw):
    toks = [red(t) for t in re.findall(r"[a-z0-9]+", lat(raw))]
    toks = [t for t in toks if t]
    core = [t for t in toks if t not in {red(p) for p in PATRON}]
    if not core:
        return None
    g = ""
    low = lat(raw)
    if re.search(r"\b(qizi|kizi|kiza)\b", low) or re.search(r"(ovna|evna|ichna)$",
                                                            core[-1] if core else ""):
        g = "f"
    elif re.search(r"\b(o?g?li|ugli)\b", low) or re.search(r"(vich|ich)$",
                                                           core[-1] if core else ""):
        g = "m"
    sur = re.sub(r"(ova|eva|yeva)$", "#f", core[0])
    sur = re.sub(r"(ov|ev|yev)$", "#", sur)
    first = core[1] if len(core) > 1 else ""
    patr = PATRON_SUF.sub("", core[2]) if len(core) > 2 else ""
    return {"full": " ".join(toks), "stem": f"{sur}|{first}|{patr}",
            "sur": sur, "s": f"{sur} {first} {patr}".strip(), "g": g}


def addr_tokens(raw):
    toks = re.findall(r"[a-z0-9]+", lat(raw))
    out = set()
    for t in toks:
        r = red(t)
        if r and len(r) >= 3 and r not in {red(x) for x in ADDR_STOP} \
                and not r.isdigit():
            out.add(r)
    return out


def clean_pinfl(v):
    if v is None or nfc(v).strip() == "":
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = re.sub(r"\D", "", str(v))
    return s if 12 <= len(s) <= 16 else ""


# -------------------------------------------------- ustunlarni aniqlash
def name_like(v):
    s = nfc(v).strip()
    if not s or HEX64.match(s):
        return False
    parts = s.split()
    return len(parts) >= 2 and len(s) >= 6 and all(
        re.fullmatch(r"[A-Za-zА-Яа-яЎўҚқҒғҲҳ`'’ʼ.\-]+", p) for p in parts)


def addr_like(v):
    s = lat(v)
    return len(s) >= 12 and (" " in s) and bool(
        re.search(r"viloyat|tuman|mahalla|kucha|ko|mfy|msg|район|область|уй|dom|uy", s))


def pinfl_like(v):
    return bool(clean_pinfl(v))


def detect(path, need_pinfl):
    wb = load_workbook(path, read_only=True)
    best = None
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        head = []
        for _ in range(6):
            try:
                head.append(next(rows))
            except StopIteration:
                break
        if not head:
            continue
        # sarlavha qatori (hint so'zlari bo'yicha)
        hdr_row, cols = None, {}
        for i, row in enumerate(head):
            found = {}
            for j, v in enumerate(row):
                h = lat(v)
                if not h:
                    continue
                for role, words in HINTS.items():
                    if role not in found and any(w in h for w in words):
                        found[role] = j
            if "name" in found or "pinfl" in found:
                hdr_row, cols = i, found
                break
        # ustunlarni kontent bo'yicha to'ldirish
        sample = head[(hdr_row + 1) if hdr_row is not None else 0:]
        # ko'proq qator o'qib namuna yig'amiz
        for _ in range(300):
            try:
                sample.append(next(rows))
            except StopIteration:
                break
        ncol = max((len(r) for r in sample), default=0)

        def best_col(pred, exclude=()):
            bestj, bestshare = None, 0.0
            for j in range(ncol):
                if j in exclude:
                    continue
                vals = [r[j] for r in sample if j < len(r)]
                vals = [v for v in vals if nfc(v).strip() != ""]
                if not vals:
                    continue
                share = sum(bool(pred(v)) for v in vals) / len(vals)
                if share > bestshare:
                    bestj, bestshare = j, share
            return bestj if bestshare >= 0.5 else None

        if "name" not in cols:
            nj = best_col(name_like)
            if nj is not None:
                cols["name"] = nj
        if "address" not in cols:
            aj = best_col(addr_like, exclude={cols.get("name")})
            if aj is not None:
                cols["address"] = aj
        if need_pinfl and "pinfl" not in cols:
            pj = best_col(pinfl_like, exclude={cols.get("name"), cols.get("address")})
            if pj is not None:
                cols["pinfl"] = pj

        ok = ("name" in cols) and (not need_pinfl or "pinfl" in cols)
        score = (1 if ok else 0, ws.max_row)
        if ok and (best is None or score > best[0]):
            best = (score, ws.title, hdr_row, dict(cols))
    wb.close()
    if not best:
        return None
    return {"sheet": best[1], "hdr": best[2], "cols": best[3]}


# --------------------------------------------------------- manba indeksi
class Index:
    def __init__(self):
        self.by_full = defaultdict(list)
        self.by_stem = defaultdict(list)
        self.by_sur = defaultdict(list)
        self.rows = []     # (pinfl, addr_set, gender, name, s)

    def add(self, name, pinfl, addr):
        k = name_keys(name)
        if not k or not pinfl:
            return
        i = len(self.rows)
        self.rows.append((pinfl, addr_tokens(addr), k["g"], nfc(name).strip(), k["s"]))
        self.by_full[k["full"]].append(i)
        self.by_stem[k["stem"]].append(i)
        self.by_sur[k["sur"]].append(i)

    def _dedupe(self, idxs):
        by_pinfl = {}
        for i in idxs:
            by_pinfl.setdefault(self.rows[i][0], i)
        return list(by_pinfl.values())

    def _pick_by_addr(self, idxs, addr):
        if not USE_ADDRESS or not addr:
            return None
        scored = sorted(((len(self.rows[i][1] & addr), i) for i in idxs),
                        reverse=True)
        if scored and scored[0][0] >= 1 and (
                len(scored) == 1 or scored[0][0] > scored[1][0]):
            return scored[0][1]
        return None

    def lookup(self, name, addr):
        """-> (pinfl|None, holat, daraja, nomzod_idxlar)"""
        k = name_keys(name)
        if not k:
            return None, "ism bo'sh", "bosh", []
        aset = addr_tokens(addr)
        for store, label, exact in ((self.by_full, "aniq", True),
                                    (self.by_stem, "aniq (yozuv farqi)", True)):
            idxs = self._dedupe(store.get(k["full"] if exact and store is self.by_full
                                          else k["stem"], []))
            if not idxs:
                continue
            if len(idxs) == 1:
                return self.rows[idxs[0]][0], label, "aniq", idxs
            pick = self._pick_by_addr(idxs, aset)
            if pick is not None:
                return self.rows[pick][0], f"{label} + manzil bo'yicha", "taxminiy", idxs
            return None, f"noaniq: {len(idxs)} ta bir xil ism", "noaniq", idxs
        # fuzzy - faqat familiyadosh nomzodlar ichida
        cands = self.by_sur.get(k["sur"], [])
        if len(k["s"]) >= 8 and cands:
            sm = SequenceMatcher(None)
            sm.set_seq2(k["s"])
            scored = []
            for i in cands:
                if k["g"] and self.rows[i][2] and k["g"] != self.rows[i][2]:
                    continue
                sm.set_seq1(self.rows[i][4])
                if sm.real_quick_ratio() < FUZZY_THRESHOLD:
                    continue
                r = sm.ratio()
                if r >= FUZZY_THRESHOLD:
                    scored.append((r, i))
            if scored:
                scored.sort(reverse=True)
                best = self._dedupe([i for _, i in scored])
                if len(best) == 1:
                    return self.rows[best[0]][0], f"imlo farqi ({scored[0][0]:.0%})", \
                        "taxminiy", best
                pick = self._pick_by_addr(best, aset)
                if pick is not None:
                    return self.rows[pick][0], "imlo farqi + manzil", "taxminiy", best
                return None, "noaniq: o'xshash ismlar", "noaniq", best
        return None, "topilmadi", "topilmadi", []


# ------------------------------------------------------------- bezak
FILL = {"aniq": PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA"),
        "taxminiy": PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC"),
        "noaniq": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD"),
        "topilmadi": PatternFill("solid", start_color="F8CBAD", end_color="F8CBAD")}
LABEL = {"aniq": "✔ Topildi", "taxminiy": "≈ Taxminiy", "noaniq": "? Noaniq",
         "topilmadi": "✘ Topilmadi", "bosh": "—"}


def resolve(name, title):
    p = name if os.path.isabs(name) else os.path.join(BASE_DIR, name)
    if os.path.exists(p):
        return p
    key = re.sub(r"[^a-z0-9]", "", name.lower())
    for f in sorted(os.listdir(BASE_DIR)):
        if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$") \
                and re.sub(r"[^a-z0-9]", "", f.lower()) == key:
            return os.path.join(BASE_DIR, f)
    files = [f for f in sorted(os.listdir(BASE_DIR))
             if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~$")]
    print(f"'{name}' topilmadi. {title}")
    for i, f in enumerate(files, 1):
        print(f"  {i}. {f}")
    try:
        return os.path.join(BASE_DIR, files[int(input("Raqamini kiriting: ")) - 1])
    except (ValueError, IndexError, EOFError):
        return None


def col_letter_to_idx(letter):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter) - 1


def main():
    src = resolve(SOURCE_FILE, "Manba (PINFL) faylni tanlang:")
    tgt = resolve(TARGET_FILE, "Natija (hashlangan) faylni tanlang:")
    if not src or not tgt:
        print("Fayllar tanlanmadi.")
        return
    print("Manba :", os.path.basename(src))
    print("Natija:", os.path.basename(tgt))
    t0 = time.time()

    print("Manba fayl aniqlanmoqda...")
    sd = detect(src, need_pinfl=True)
    if not sd:
        print("Manba faylda PINFL va ism ustunlari topilmadi.")
        return
    print(f"  Manba: [{sd['sheet']}] ism={get_column_letter(sd['cols']['name'] + 1)}, "
          f"PINFL={get_column_letter(sd['cols']['pinfl'] + 1)}, "
          f"manzil={get_column_letter(sd['cols'].get('address', -1) + 1) if 'address' in sd['cols'] else '-'}")

    print("Manba o'qilmoqda (katta bo'lsa biroz kutiladi)...")
    idx = Index()
    wb_s = load_workbook(src, read_only=True)
    ws_s = wb_s[sd["sheet"]]
    ci = sd["cols"]
    rows = ws_s.iter_rows(values_only=True)
    for _ in range((sd["hdr"] or -1) + 1):
        next(rows, None)
    get = lambda row, key: row[ci[key]] if key in ci and ci[key] < len(row) else None
    n = 0
    for row in rows:
        idx.add(get(row, "name"), clean_pinfl(get(row, "pinfl")), get(row, "address"))
        n += 1
    wb_s.close()
    print(f"  Manba: {n} qator, {len(idx.rows)} ta PINFL li yozuv")

    # natija fayl
    print("Natija fayl o'qilmoqda...")
    wb = load_workbook(tgt)
    ws = wb.active
    if TARGET_NAME_COL:
        c_name0 = col_letter_to_idx(TARGET_NAME_COL)
        c_addr0 = col_letter_to_idx(TARGET_ADDR_COL) if TARGET_ADDR_COL else None
        hdr_t = 0
    else:
        td = detect(tgt, need_pinfl=False)
        if not td:
            print("Natija faylda ism ustuni topilmadi. TARGET_NAME_COL bering.")
            return
        c_name0 = td["cols"]["name"]
        c_addr0 = td["cols"].get("address")
        hdr_t = td["hdr"] or 0
    c_name, c_addr = c_name0 + 1, (c_addr0 + 1 if c_addr0 is not None else None)
    print(f"  Natija: ism={get_column_letter(c_name)}, "
          f"manzil={get_column_letter(c_addr) if c_addr else '-'} | "
          f"sarlavha qatori: {hdr_t if hdr_t else 'yo‘q'}")

    last = max((c.column for row in ws.iter_rows(min_row=1, max_row=1) for c in row
                if c.value not in (None, "")), default=ws.max_column)
    c_pinfl, c_stat = last + 1, last + 2
    start = (hdr_t + 1) if hdr_t else 1
    if hdr_t:
        ws.cell(hdr_t, c_pinfl).value = "PINFL (2200 dan)"
        ws.cell(hdr_t, c_stat).value = "Moslik"

    stats = Counter()
    review = []
    for r in range(start, ws.max_row + 1):
        name = ws.cell(r, c_name).value
        if name is None or nfc(name).strip() == "":
            continue
        addr = ws.cell(r, c_addr).value if c_addr else None
        pinfl, holat, level, cands = idx.lookup(str(name), addr)
        stats[level] += 1
        if pinfl:
            cell = ws.cell(r, c_pinfl)
            cell.value = pinfl
            cell.number_format = "@"
            cell.alignment = Alignment(horizontal="center")
        ws.cell(r, c_stat).value = LABEL.get(level, level)
        if HIGHLIGHT and level in FILL:
            ws.cell(r, c_pinfl).fill = FILL[level]
            ws.cell(r, c_stat).fill = FILL[level]
        if level != "aniq":
            cand_txt = "\n".join(
                f"{idx.rows[i][0]} | {idx.rows[i][3]}" for i in cands[:6])
            review.append([r, nfc(name).strip(), nfc(addr).strip() if addr else "",
                           LABEL.get(level, level), holat, cand_txt])

    # Tekshirish varag'i
    if "Tekshirish" in wb.sheetnames:
        del wb["Tekshirish"]
    sr = wb.create_sheet("Tekshirish")
    head = ["Qator №", "Ism (natija fayl)", "Manzil", "Moslik", "Izoh",
            "Nomzodlar (PINFL | ism, manba fayl)"]
    sr.append(head)
    for c in sr[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in review:
        sr.append(row)
        rr = sr.max_row
        for c in range(1, 7):
            sr.cell(rr, c).alignment = Alignment(vertical="top", wrap_text=c in (5, 6))
        if row[3] in LABEL.values():
            for key, lab in LABEL.items():
                if lab == row[3] and key in FILL:
                    sr.cell(rr, 4).fill = FILL[key]
    for i, w in enumerate([9, 34, 40, 13, 32, 60], 1):
        sr.column_dimensions[get_column_letter(i)].width = w
    sr.freeze_panes = "A2"
    sr.auto_filter.ref = sr.dimensions

    out = os.path.join(BASE_DIR, OUTPUT_NAME)
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(BASE_DIR, f"Hashlangan_PINFL_{datetime.now():%H%M%S}.xlsx")
        print("Natija fayli ochiq ekan, boshqa nomga saqlanmoqda.")
        wb.save(out)

    tot = sum(v for k, v in stats.items() if k != "bosh")
    ok = stats["aniq"] + stats["taxminiy"]
    print(f"Jami: {tot} | topildi: {stats['aniq']} | taxminiy: {stats['taxminiy']} | "
          f"noaniq: {stats['noaniq']} | topilmadi: {stats['topilmadi']} | "
          f"qamrov: {ok / max(1, tot):.1%}")
    print(f"PINFL ustuni: {get_column_letter(c_pinfl)} | Vaqt: {time.time() - t0:.1f} s")
    print("Saqlandi:", out)
    print("  Yashil=aniq, sariq=taxminiy (tekshiring), qizil=noaniq/topilmadi.")
    print("  \"Tekshirish\" varag'ida taxminiy/noaniq/topilmagan qatorlar.")


if __name__ == "__main__":
    try:
        main()
    except PermissionError as e:
        print("Faylga kirib bo'lmadi. Excelda ochiq bo'lsa yoping:", e)
    except Exception:
        print("Kutilmagan xato:")
        traceback.print_exc()
