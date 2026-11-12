# -*- coding: utf-8 -*-
"""
ЯТТ реестридан FAOLIYAT YURITAYOTGANLARNI ajratib oladi va "Яшаш манзили" ustunini
alohida ustunlarga bo'ladi:  Viloyat | Tuman/Shahar | QFY | MFY | Ko'cha | Uy

Misol:
  Surxondaryo viloyati Muzrabot tumani МУЗРАБОТ ҚФЙ, АЛПОМИШ МФЙ, Ғ.ҒУЛОМ, uy:.
    -> Surxondaryo viloyati | Muzrabot tumani | "MUZRABOT" QFY | "ALPOMISH" MFY | G`.G`ULOM |
  Surxondaryo viloyati Termiz shahar ТЕРМИЗ Ш., БОҒЗОР МФЙ, ИСЛОМ КАРИМОВ КЎЧАСИ, uy:95
    -> Surxondaryo viloyati | Termiz shahri |  | "BOG`ZOR" MFY | ISLOM KARIMOV KO`CHASI | 95
  ... Термезский район, Намуна МСГ, ул. Гулзор, дом 87
    -> ... | "NAMUNA" MFY | GULZOR | 87

Nimalarni hisobga oladi:
  - МФЙ / MFY / m.f.y. / МСГ / mahallasi  -> MFY;   ҚФЙ / QFY / КФЙ / ССГ -> QFY
  - takrorlar ("ХЎЖАНҚОН МФЙ, ... ХЎЖАНҚОН МФЙ") bitta qilinadi
  - "ТЕРМИЗ Ш.", "Сурхандарьинская область", "Термез г.", "Термезский район" kabi ortiqcha qismlar tashlanadi
  - uy: 116 xonadon:5 / 48А-уй / дом 13 / 1 Квартира -> Uy ustuniga
  - kirill yozuvi o'zbek lotiniga o'giriladi (sozlamada o'chirish mumkin)

Natija: <nom>_faol.xlsx
  1-varaq — faqat faoliyat yuritayotgan YaTTlar (barcha asl ustunlar + 6 ta yangi ustun)
  "MFY kesimida" — tuman va MFY bo'yicha soni
  "Ajratilmadi" — MFY topilmagan qatorlar (qo'lda tekshirish uchun)

Ishga tushirish:
  py -m pip install openpyxl
  py yatt_manzil.py
"""

import os
import re
import sys
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ================== ASOSIY SOZLAMALAR ==================
INPUT_FILE = "ЯТТ реестири (2) (2).xlsx"     # topilmasa, ro'yxatdan tanlatadi
ADDRESS_HEADER = "яшаш манзили"               # bo'linadigan ustun
STATUS_HEADER = "ятт холати"                  # holat ustuni
ONLY_ACTIVE = True                            # faqat "фаолият юритаяпти"
TRANSLIT = True                               # kirillni lotinga o'girish
APOSTROPHE = "`"                              # o`, g` belgisi
# =======================================================

NEW_COLS = ["Viloyat", "Tuman/Shahar", "QFY", "MFY", "Ko'cha", "Uy"]
HEAD_FILL = PatternFill("solid", fgColor="C6E0B4")

# ---------------- Kirill -> o'zbek lotin ----------------
_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "yo", "ж": "j", "з": "z", "и": "i",
    "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t",
    "у": "u", "ф": "f", "х": "x", "ц": "s", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": APOSTROPHE, "ы": "i",
    "ь": "", "э": "e", "ю": "yu", "я": "ya", "ў": "o" + APOSTROPHE, "қ": "q", "ғ": "g" + APOSTROPHE,
    "ҳ": "h",
}
CYR_RE = re.compile(r"[а-яёўқғҳ]", re.I)


def to_latin(s):
    if not TRANSLIT or not CYR_RE.search(s):
        return s
    out = []
    for i, ch in enumerate(s):
        low = ch.lower()
        if low == "е" and (i == 0 or not s[i - 1].isalpha()):
            t = "ye"                                   # so'z boshida Е -> YE
        elif low in _MAP:
            t = _MAP[low]
        else:
            out.append(ch)
            continue
        if ch.isupper():
            nxt = s[i + 1] if i + 1 < len(s) else ""
            t = t.upper() if (nxt.isupper() or not nxt.isalpha()) else t.capitalize()
        out.append(t)
    return "".join(out)


def unify_apos(s):
    return re.sub(r"(?<=\w)['‘’ʻʼ´`](?=\w)", APOSTROPHE, s)


# ---------------- Manzilni bo'lish ----------------
# 1-qism: "Surxondaryo viloyati Termiz shahar ..." / "Toshkent shahri Olmazor tumani ..."
PREFIX_RE = re.compile(
    r"^\s*(?P<vil>.+?\s(?:viloyati|вилояти|Respublikasi|Республикаси)|Toshkent\s+shahri|Тошкент\s+шаҳри)\s+"
    r"(?P<tum>.+?\s(?:tumani|туман(?:и)?|shahri|shahar|шаҳри|шахри|шаҳар|шахар))(?=\s|,|$)\s*,?\s*(?P<rest>.*)$",
    re.I)

MFY_RE = re.compile(r"(?<![\w])(МФЙ|MFY|М\.Ф\.Й\.?|M\.F\.Y\.?|МСГ|махалласи|маҳалласи|mahallasi|"
                    r"махалла|маҳалла|mahalla)(?![\w])", re.I)
QFY_RE = re.compile(r"(?<![\w])(ҚФЙ|КФЙ|QFY|Қ\.Ф\.Й\.?|Q\.F\.Y\.?|ССГ|ШФЙ|SHFY)(?![\w])", re.I)

# tashlab yuboriladigan qismlar (tuman/viloyat takrori)
JUNK_RE = re.compile(
    r"^(\S+\s+)?(ш|sh|г|g|т|t|шаҳар|шахар|shahar|shahri|шаҳри|шахри|район|tumani|тумани|туман|область|"
    r"вилояти|viloyati|обл)\.?$"
    r"|^(узбекистан|ўзбекистон(\s+республикаси)?|o.zbekiston(\s+respublikasi)?)$", re.I)

HOUSE_RE = re.compile(
    r"^\s*(uy|уй|дом|д\.|xonadon|хонадон|кв\.?|квартира|kv\.?)\s*[:.]?\s*"
    r"|\b\d+\s*[\"«]?\s*[а-яa-z]?[\"»]?\s*-?\s*(uy|уй|xonadon|хонадон|дом|квартира|кв)\b"
    r"|\b(дом|квартира|xonadon|хонадон)\b", re.I)

STREET_PREFIX = re.compile(r"^(ул\.|улица|пр\.|проспект|пер\.|переулок)\s*", re.I)


def tidy(s):
    s = re.sub(r"\s+", " ", s).strip(" ,.;:")
    return s


def style_org(part, kind):
    """'ҚОРАСУВ ҚФЙ' -> '"QORASUV" QFY'"""
    rx = MFY_RE if kind == "MFY" else QFY_RE
    name = rx.sub(" ", part)
    name = re.sub(r"[\"«»“”']", " ", name)
    name = re.sub(r"\bm\.?f\.?y\.?\b|\bмфй\b", " ", name, flags=re.I)
    name = tidy(unify_apos(to_latin(name)))
    return f'"{name.upper()}" {kind}' if name else ""


def house_text(part):
    p = to_latin(part)
    # xonadon:5 / kv.12  -> 5-xonadon
    p = re.sub(r"\b(xonadon|kvartira|kv)\s*[:.]\s*(\S+)", r"\2-xonadon", p, flags=re.I)
    # 13-xonadon / 1 Kvartira / 18-хонадон -> 13-xonadon
    p = re.sub(r"(\d\S*?)\s*-?\s*\b(xonadon|kvartira|kv)\b\.?", r"\1-xonadon", p, flags=re.I)
    # 39-uy / uy:95 / dom 13 / 31 Dom -> 39 / 95 / 13 / 31
    p = re.sub(r"\s*-?\s*\b(uy|dom|d)\b\s*[:.]?", " ", p, flags=re.I)
    p = tidy(p)
    p = re.sub(r"\s+(?=\S+-xonadon)", ", ", p)
    return "" if p in ("0", ".", "-") else p


def parse_address(addr):
    res = dict.fromkeys(NEW_COLS, "")
    if addr is None or not str(addr).strip():
        return res
    s = str(addr).replace(" ", " ").strip()

    m = PREFIX_RE.match(s)
    if m:
        res["Viloyat"] = tidy(m.group("vil"))
        tum = tidy(m.group("tum"))
        tum = re.sub(r"\bshahar$", "shahri", tum, flags=re.I)
        res["Tuman/Shahar"] = to_latin(tum)
        rest = m.group("rest")
    else:
        rest = s

    parts = [tidy(p) for p in rest.split(",")]
    mfy, qfy, street, house = [], [], [], []

    for p in parts:
        if not p:
            continue
        # "uy:116 xonadon:5" — bitta qismda ikkalasi
        if re.match(r"^(uy|xonadon)\s*:", p, re.I):
            h = house_text(p)
            if h and h not in house:
                house.append(h)
            continue
        if JUNK_RE.match(p) and not MFY_RE.search(p) and not QFY_RE.search(p):
            continue
        if QFY_RE.search(p):
            v = style_org(p, "QFY")
            if v and v not in qfy:
                qfy.append(v)
            continue
        if MFY_RE.search(p):
            # "Намуна МСГ" yoki "Темир йулчи МФЙ Шероб..." — MFY dan keyingi qismi ko'cha bo'lishi mumkin
            mm = MFY_RE.search(p)
            tail = tidy(p[mm.end():])
            v = style_org(p[:mm.end()], "MFY")
            if v and v not in mfy:
                mfy.append(v)
            if tail:
                street.append(tail)
            continue
        if HOUSE_RE.search(p) and not re.search(r"(кўчаси|кучаси|ko.chasi|kuchasi|бурилиш|tor ko)", p, re.I):
            h = house_text(p)
            if h and h not in house:
                house.append(h)
            continue
        # "Исо Ат Термизий14-бурилиш" -> "Исо Ат Термизий 14-бурилиш"
        p = re.sub(r"(?<=[^\W\d])(?=\d)", " ", p)
        p = STREET_PREFIX.sub("", p)
        if p and p not in street:
            street.append(p)

    res["QFY"] = ", ".join(qfy)
    res["MFY"] = ", ".join(mfy)
    res["Ko'cha"] = tidy(unify_apos(to_latin(", ".join(street))))
    res["Uy"] = ", ".join(x for x in house if x)
    return res


# ---------------- Excel ----------------
def pick(path):
    if path and os.path.exists(path):
        return path
    files = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xlsm"))
             and not f.startswith("~$") and "_faol" not in f]
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


def norm_h(v):
    return re.sub(r"\s+", " ", str(v or "")).strip().lower()


def find_header(ws):
    for r in range(1, min(ws.max_row, 20) + 1):
        row = {norm_h(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
        if ADDRESS_HEADER in row:
            return r, row
    return None, None


def is_active(v):
    s = norm_h(v)
    if not s:
        return False
    if "тўхтат" in s or "тухтат" in s or "to'xtat" in s or "toxtat" in s or "тугат" in s:
        return False
    return "юритаяпти" in s or "yuritayapti" in s or "yuritmoqda" in s or s in ("faol", "фаол")


def copy_style(src, dst):
    if src.has_style:
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.number_format = copy(src.alignment), src.number_format


def main():
    inp = pick(sys.argv[1] if len(sys.argv) > 1 else INPUT_FILE)
    out = os.path.splitext(inp)[0] + "_faol.xlsx"
    print(f"Ochilmoqda: {inp} (katta fayl bo'lsa biroz kutasiz)")
    src_wb = load_workbook(inp, data_only=True)

    ws = None
    for w in src_wb.worksheets:
        hdr, cols = find_header(w)
        if hdr:
            ws = w
            break
    if not ws:
        sys.exit(f"'{ADDRESS_HEADER}' ustuni topilmadi.")
    a_col = cols[ADDRESS_HEADER]
    st_col = cols.get(STATUS_HEADER)
    if ONLY_ACTIVE and not st_col:
        sys.exit(f"'{STATUS_HEADER}' ustuni topilmadi.")
    tuman_code_col = next((c for h, c in cols.items() if "туман коди" in h or "tuman kodi" in h), None)
    print(f"  varaq: {ws.title}, sarlavha qatori: {hdr}, manzil ustuni: {get_column_letter(a_col)}")

    wb = Workbook()
    o = wb.active
    o.title = (ws.title or "YaTT")[:31]
    n_new = len(NEW_COLS)
    maxc = ws.max_column

    def dst_col(c):
        return c if c <= a_col else c + n_new

    # sarlavhadan oldingi qatorlar (nomlanish) va sarlavha
    for r in range(1, hdr + 1):
        for c in range(1, maxc + 1):
            s = ws.cell(r, c)
            d = o.cell(r, dst_col(c), s.value)
            copy_style(s, d)
        o.row_dimensions[r].height = ws.row_dimensions[r].height
    for i, name in enumerate(NEW_COLS):
        d = o.cell(hdr, a_col + 1 + i, name)
        copy_style(ws.cell(hdr, a_col), d)
        d.fill = HEAD_FILL
        d.font = Font(bold=True)
        d.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ustun kengliklari
    for c in range(1, maxc + 1):
        w = ws.column_dimensions[get_column_letter(c)].width
        if w:
            o.column_dimensions[get_column_letter(dst_col(c))].width = w
    for i, w in enumerate([20, 18, 22, 26, 34, 14]):
        o.column_dimensions[get_column_letter(a_col + 1 + i)].width = w

    total = active = 0
    out_r = hdr
    kesim, bad = {}, []
    for r in range(hdr + 1, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, maxc + 1)]
        if all(v in (None, "") for v in vals):
            continue
        total += 1
        if ONLY_ACTIVE and not is_active(vals[st_col - 1]):
            continue
        active += 1
        out_r += 1
        for c in range(1, maxc + 1):
            s = ws.cell(r, c)
            d = o.cell(out_r, dst_col(c), s.value)
            copy_style(s, d)
        p = parse_address(vals[a_col - 1])
        for i, k in enumerate(NEW_COLS):
            d = o.cell(out_r, a_col + 1 + i, p[k] or None)
            copy_style(ws.cell(r, a_col), d)
        key = (p["Tuman/Shahar"], p["MFY"] or "(MFY yo'q)")
        kesim[key] = kesim.get(key, 0) + 1
        if not p["MFY"]:
            bad.append([out_r, vals[tuman_code_col - 1] if tuman_code_col else "", vals[a_col - 1],
                        p["Tuman/Shahar"], p["QFY"], p["Ko'cha"]])
        if active % 2000 == 0:
            print(f"  {active} ta faol qator...")

    o.freeze_panes = o.cell(hdr + 1, 1)
    o.auto_filter.ref = f"A{hdr}:{get_column_letter(maxc + n_new)}{out_r}"

    k = wb.create_sheet("MFY kesimida")
    k.append(["Tuman/Shahar", "MFY", "Faol YaTT soni"])
    for (t, m), n in sorted(kesim.items()):
        k.append([t, m, n])
    k.append(["", "JAMI", active])
    for x in k[1]:
        x.font = Font(bold=True)
    k.column_dimensions["A"].width, k.column_dimensions["B"].width, k.column_dimensions["C"].width = 22, 32, 15
    k.auto_filter.ref = f"A1:C{k.max_row - 1}"

    b = wb.create_sheet("Ajratilmadi")
    b.append(["Qator (1-varaqda)", "Tuman kodi", "Яшаш манзили (asl)", "Tuman/Shahar", "QFY", "Ko'cha"])
    for x in b[1]:
        x.font = Font(bold=True)
    for row in bad:
        b.append(row)
    for col, w in zip("ABCDEF", [10, 10, 90, 20, 22, 34]):
        b.column_dimensions[col].width = w

    try:
        wb.save(out)
    except PermissionError:
        sys.exit(f"'{out}' ochiq turibdi — yopib, qayta ishga tushiring.")
    print(f"\nSaqlandi: {out}")
    print(f"  Jami: {total}, faoliyat yuritayotgan: {active}")
    print(f"  MFY ajratildi: {active - len(bad)}, ajratilmadi: {len(bad)} ('Ajratilmadi' varag'i)")


if __name__ == "__main__":
    main()
