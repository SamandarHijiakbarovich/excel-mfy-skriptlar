# -*- coding: utf-8 -*-
"""
Excel'dagi manzillarni xaritada qidirib, har bir qatorga aniq Google Maps linkini qo'yadi.

Ishlash tartibi (har bir qator uchun):
  1. Manzil tozalanadi (xonadon/kv olib tashlanadi, `‘’ belgilar bir xil qilinadi, "sh." -> "shahri").
  2. Bir nechta variant bilan qidiriladi (aniqdan umumiyga):
       uy darajasi   -> to'liq manzil
       ko'cha        -> uy raqamisiz manzil
       mahalla       -> MFY nomi + Tuman/Shahar + Viloyat
       shahar/tuman  -> Tuman/Shahar + Viloyat
  3. Har bir natija Surxondaryo hududi ichida ekanligi tekshiriladi (tashqaridagilar rad etiladi).
  4. Eng aniq natija tanlanadi va "Aniqlik" ustuniga yoziladi.
  5. Hech narsa topilmasa — oddiy qidiruv linki qo'yiladi va "topilmadi" deb belgilanadi.

Natijalar keshlanadi (geocache.json) — skript to'xtab qolsa, qayta ishga tushirganda
oldin qidirilganlar qayta so'ralmaydi. Har 50 qatorda fayl saqlanadi.

Provayderlar:
  google    — eng yaxshi, API kalit kerak (Geocoding API yoqilgan bo'lishi kerak)
  yandex    — O'zbekistonda ko'pincha aniqroq, API kalit kerak (Geocoder API)
  nominatim — bepul (OpenStreetMap), kalitsiz, lekin sekin (1 so'rov/sek) va kamroq aniq

Misollar:
  pip install openpyxl requests
  python xarita_link.py -i "1722 ЮРИДИК+ 401 (1).xlsx" -p google -k AIza... --limit 20
  python xarita_link.py -i "1722 ЮРИДИК+ 401 (1).xlsx" -p yandex -k 1234-abcd...
  python xarita_link.py -i "1722 ЮРИДИК+ 401 (1).xlsx" -p nominatim
"""

import argparse
import json
import os
import re
import sys
import time
from urllib.parse import quote_plus

import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

# ================== ASOSIY SOZLAMALAR (shu yerni o'zgartiring) ==================
# Shunda "py main.py" deb argumentsiz ishga tushirish mumkin.
DEFAULT_INPUT = "1722 ЮРИДИК+ 401 (1).xlsx"   # topilmasa, papkadagi .xlsx fayllardan tanlatadi
DEFAULT_PROVIDER = "nominatim"                 # "google", "yandex" yoki "nominatim" (bepul)
DEFAULT_KEY = ""                               # google/yandex uchun API kalit
DEFAULT_LIMIT = None                           # sinov uchun masalan 20; hammasi uchun None
# ================================================================================

# ---------------- Sozlamalar ----------------
# Surxondaryo viloyati chegarasi (taxminiy): natija shu oraliqdan tashqarida bo'lsa rad etiladi
BBOX = {"lat_min": 37.0, "lat_max": 39.1, "lon_min": 66.4, "lon_max": 68.7}

HEADER_NAMES = {
    "manzil": ["manzil"],
    "viloyat": ["viloyat nomi", "viloyat"],
    "tuman": ["tuman/shahar", "tuman", "shahar"],
    "mfy": ["mfy nomi", "mahalla"],
}

OUT_COLS = ["Xarita link", "Koordinata", "Aniqlik", "Topilgan so'rov", "Manba"]

LEVEL_NAMES = {4: "uy", 3: "ko'cha", 2: "mahalla", 1: "shahar/tuman", 0: "topilmadi"}
LEVEL_COLORS = {4: "C6EFCE", 3: "E2EFDA", 2: "FFEB9C", 1: "FCE4D6", 0: "FFC7CE"}

CACHE_FILE = "geocache.json"
SAVE_EVERY = 50
USER_AGENT = "xarita-link-script/1.0 (excel geocoding)"


# ---------------- Manzilni tozalash ----------------
def norm(s):
    if s is None:
        return ""
    s = str(s)
    for ch in "`‘’ʻʼ´":
        s = s.replace(ch, "'")
    s = re.sub(r"\s+", " ", s).strip(" ,.;")
    return s


def expand_abbr(s):
    s = re.sub(r"\b(\w+)\s+sh\.(?=\s|,|$)", r"\1 shahri", s, flags=re.I)
    s = re.sub(r"\b(\w+)\s+t\.(?=\s|,|$)", r"\1 tumani", s, flags=re.I)
    s = re.sub(r"\bko'ch\.(?=\s|,|$)", "ko'chasi", s, flags=re.I)
    s = re.sub(r"\bm\.?f\.?y\.?\s*", "", s, flags=re.I)
    # "Termiz shahri Eski shaxar mahallasi" -> "Termiz shahri, Eski shaxar mahallasi"
    s = re.sub(r"\b(shahri|tumani)\s+(?=[^,\s])", r"\1, ", s, flags=re.I)
    return s


RE_FLAT = re.compile(r"(xonadon|xona|kv\.?|kvartira|\bxon\b)", re.I)
RE_HOUSE = re.compile(r"(\d+\s*[a-zA-Z]?\s*-?\s*uy\b|\buy\s*\d+|\bdom\b)", re.I)


def split_parts(addr):
    return [p.strip() for p in addr.split(",") if p.strip()]


def build_queries(row):
    """Aniqdan umumiyga qarab so'rovlar ro'yxati: [(daraja, so'rov), ...]"""
    viloyat = norm(row.get("viloyat")) or "Surxondaryo viloyati"
    tuman = expand_abbr(norm(row.get("tuman")))
    mfy = expand_abbr(norm(row.get("mfy")))
    addr = expand_abbr(norm(row.get("manzil")))

    parts = [p for p in split_parts(addr) if not RE_FLAT.search(p)]
    queries = []

    if parts:
        full = ", ".join(parts)
        queries.append((4, full))
        no_house = [p for p in parts if not RE_HOUSE.search(p)]
        if no_house and no_house != parts:
            queries.append((3, ", ".join(no_house)))

    if mfy:
        m = mfy if re.search(r"mahalla", mfy, re.I) else f"{mfy} mahallasi"
        queries.append((2, f"{m}, {tuman}, {viloyat}" if tuman else f"{m}, {viloyat}"))

    if tuman:
        queries.append((1, f"{tuman}, {viloyat}"))

    # takrorlarni olib tashlash
    seen, uniq = set(), []
    for lvl, q in queries:
        k = q.lower()
        if k not in seen:
            seen.add(k)
            uniq.append((lvl, q + ", O'zbekiston"))
    return uniq


def in_bbox(lat, lon):
    return BBOX["lat_min"] <= lat <= BBOX["lat_max"] and BBOX["lon_min"] <= lon <= BBOX["lon_max"]


# ---------------- Provayderlar ----------------
class Geocoder:
    def __init__(self, provider, key):
        self.provider = provider
        self.key = key
        self.s = requests.Session()
        self.s.headers["User-Agent"] = USER_AGENT
        self.delay = 1.1 if provider == "nominatim" else 0.05
        self._last = 0.0

    def _wait(self):
        dt = time.time() - self._last
        if dt < self.delay:
            time.sleep(self.delay - dt)
        self._last = time.time()

    def _get(self, url, params):
        for attempt in range(4):
            self._wait()
            try:
                r = self.s.get(url, params=params, timeout=20)
                if r.status_code == 429 or r.status_code >= 500:
                    time.sleep(2 ** attempt * 2)
                    continue
                r.raise_for_status()
                return r.json()
            except requests.RequestException as e:
                if attempt == 3:
                    print(f"   ! tarmoq xatosi: {e}")
                    return None
                time.sleep(2 ** attempt)
        return None

    def geocode(self, q):
        """-> dict(lat, lon, level) yoki None"""
        return getattr(self, "_" + self.provider)(q)

    # Google: location_type va types bo'yicha aniqlik
    def _google(self, q):
        j = self._get("https://maps.googleapis.com/maps/api/geocode/json", {
            "address": q, "key": self.key, "region": "uz", "language": "uz",
            "components": "country:UZ",
            "bounds": f"{BBOX['lat_min']},{BBOX['lon_min']}|{BBOX['lat_max']},{BBOX['lon_max']}",
        })
        if not j:
            return None
        st = j.get("status")
        if st in ("REQUEST_DENIED", "OVER_DAILY_LIMIT", "INVALID_REQUEST"):
            sys.exit(f"Google xatosi: {st} — {j.get('error_message', '')}")
        if st != "OK":
            return None
        for res in j["results"]:
            loc = res["geometry"]["location"]
            if not in_bbox(loc["lat"], loc["lng"]):
                continue
            t = set(res.get("types", []))
            lt = res["geometry"].get("location_type")
            if lt == "ROOFTOP" or t & {"street_address", "premise", "subpremise"}:
                lvl = 4
            elif lt == "RANGE_INTERPOLATED" or "route" in t:
                lvl = 3
            elif t & {"neighborhood", "sublocality", "sublocality_level_1", "sublocality_level_2"}:
                lvl = 2
            else:
                lvl = 1
            if res.get("partial_match") and lvl > 1:
                lvl -= 1
            return {"lat": loc["lat"], "lon": loc["lng"], "level": lvl}
        return None

    # Yandex: kind va precision bo'yicha
    def _yandex(self, q):
        j = self._get("https://geocode-maps.yandex.ru/1.x/", {
            "apikey": self.key, "geocode": q, "format": "json", "lang": "ru_RU",
            "bbox": f"{BBOX['lon_min']},{BBOX['lat_min']}~{BBOX['lon_max']},{BBOX['lat_max']}",
            "rspn": 1, "results": 5,
        })
        if not j:
            return None
        if "statusCode" in j and j["statusCode"] in (401, 403):
            sys.exit(f"Yandex xatosi: {j.get('message')}")
        try:
            members = j["response"]["GeoObjectCollection"]["featureMember"]
        except KeyError:
            return None
        for m in members:
            g = m["GeoObject"]
            lon, lat = map(float, g["Point"]["pos"].split())
            if not in_bbox(lat, lon):
                continue
            meta = g["metaDataProperty"]["GeocoderMetaData"]
            kind, prec = meta.get("kind"), meta.get("precision")
            if kind == "house" and prec in ("exact", "number"):
                lvl = 4
            elif kind in ("house", "street"):
                lvl = 3
            elif kind == "district":
                lvl = 2
            else:
                lvl = 1
            return {"lat": lat, "lon": lon, "level": lvl}
        return None

    # Nominatim (OpenStreetMap): addresstype bo'yicha
    def _nominatim(self, q):
        j = self._get("https://nominatim.openstreetmap.org/search", {
            "q": q, "format": "jsonv2", "countrycodes": "uz", "limit": 5,
            "viewbox": f"{BBOX['lon_min']},{BBOX['lat_max']},{BBOX['lon_max']},{BBOX['lat_min']}",
            "bounded": 1, "accept-language": "uz",
        })
        if not j:
            return None
        for res in j:
            lat, lon = float(res["lat"]), float(res["lon"])
            if not in_bbox(lat, lon):
                continue
            at = res.get("addresstype") or res.get("type", "")
            if at in ("house", "building", "amenity", "shop", "office"):
                lvl = 4
            elif at in ("road", "street", "pedestrian", "residential"):
                lvl = 3
            elif at in ("neighbourhood", "suburb", "quarter", "hamlet", "city_district"):
                lvl = 2
            else:
                lvl = 1
            return {"lat": lat, "lon": lon, "level": lvl}
        return None


# ---------------- Qator uchun eng yaxshi natija ----------------
def resolve(row, geo, cache):
    best = None
    for q_lvl, q in build_queries(row):
        ck = f"{geo.provider}|{q.lower()}"
        if ck in cache:
            res = cache[ck]
        else:
            res = geo.geocode(q)
            cache[ck] = res
        if not res:
            continue
        lvl = min(q_lvl, res["level"])  # so'rov umumiy bo'lsa, natija undan aniq bo'lolmaydi
        if best is None or lvl > best["level"]:
            best = {**res, "level": lvl, "query": q}
        if res["level"] >= q_lvl:   # so'ralgan aniqlikda topildi — keyingisi kerak emas
            break
    return best


# ---------------- Excel ----------------
def find_header(ws):
    for r in range(1, 11):
        vals = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None:
                vals[str(v).strip().lower()] = c
        if "manzil" in vals:
            cols = {}
            for key, names in HEADER_NAMES.items():
                for n in names:
                    if n in vals:
                        cols[key] = vals[n]
                        break
            return r, cols
    sys.exit("'Manzil' ustuni birinchi 10 qatordan topilmadi.")


def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache):
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)
    os.replace(tmp, CACHE_FILE)


def main():
    ap = argparse.ArgumentParser(description="Manzillarga Google Maps link qo'yish")
    ap.add_argument("-i", "--input", default=DEFAULT_INPUT, help="Kirish Excel fayli")
    ap.add_argument("-o", "--output", help="Chiqish fayli (standart: <nom>_xarita.xlsx)")
    ap.add_argument("-p", "--provider", choices=["google", "yandex", "nominatim"], default=DEFAULT_PROVIDER)
    ap.add_argument("-k", "--key", default=DEFAULT_KEY or None, help="API kalit (google/yandex uchun)")
    ap.add_argument("-s", "--sheet", help="Varaq nomi (standart: birinchi)")
    ap.add_argument("--limit", type=int, default=DEFAULT_LIMIT, help="Faqat birinchi N qatorni ishlash (sinov uchun)")
    a = ap.parse_args()

    # Fayl nomi berilmagan yoki topilmasa — papkadagi .xlsx fayllardan tanlaymiz
    if not a.input or not os.path.exists(a.input):
        files = [f for f in os.listdir(".")
                 if f.lower().endswith(".xlsx") and not f.startswith("~$") and "_xarita" not in f]
        if not files:
            sys.exit("Papkada .xlsx fayl topilmadi. Faylni skript bilan bir papkaga qo'ying.")
        if len(files) == 1:
            a.input = files[0]
        else:
            print("Qaysi faylni ishlaymiz?")
            for n, f in enumerate(files, 1):
                print(f"  {n}. {f}")
            while True:
                ch = input("Raqamni kiriting: ").strip()
                if ch.isdigit() and 1 <= int(ch) <= len(files):
                    a.input = files[int(ch) - 1]
                    break

    if a.provider in ("google", "yandex") and not a.key:
        sys.exit(f"{a.provider} uchun -k API kalit kerak.")

    out = a.output or os.path.splitext(a.input)[0] + "_xarita.xlsx"
    print(f"Ochilmoqda: {a.input}")
    wb = load_workbook(a.input)
    ws = wb[a.sheet] if a.sheet else wb.worksheets[0]

    hdr_row, cols = find_header(ws)
    print(f"Sarlavha qatori: {hdr_row}, topilgan ustunlar: {cols}")

    # chiqish ustunlari (oldin qo'shilgan bo'lsa, o'shalarni qayta ishlatadi)
    existing = {str(ws.cell(hdr_row, c).value).strip(): c
                for c in range(1, ws.max_column + 1) if ws.cell(hdr_row, c).value}
    start = ws.max_column + 1
    out_idx = {}
    for name in OUT_COLS:
        if name in existing:
            out_idx[name] = existing[name]
        else:
            out_idx[name] = start
            start += 1
        h = ws.cell(hdr_row, out_idx[name], name)
        h.font = Font(bold=True)
    ws.column_dimensions[ws.cell(hdr_row, out_idx["Xarita link"]).column_letter].width = 20
    ws.column_dimensions[ws.cell(hdr_row, out_idx["Topilgan so'rov"]).column_letter].width = 60

    geo = Geocoder(a.provider, a.key)
    cache = load_cache()
    last = ws.max_row if not a.limit else min(ws.max_row, hdr_row + a.limit)
    total = last - hdr_row
    stats = {k: 0 for k in LEVEL_NAMES}
    t0 = time.time()

    try:
        for i, r in enumerate(range(hdr_row + 1, last + 1), 1):
            row = {k: ws.cell(r, c).value for k, c in cols.items()}
            if not norm(row.get("manzil")) and not norm(row.get("mfy")):
                continue

            best = resolve(row, geo, cache)
            if best:
                lat, lon, lvl = best["lat"], best["lon"], best["level"]
                link = f"https://www.google.com/maps?q={lat:.6f},{lon:.6f}"
                coord = f"{lat:.6f}, {lon:.6f}"
                query = best["query"]
            else:
                lvl = 0
                q = norm(row.get("manzil")) + ", O'zbekiston"
                link = "https://www.google.com/maps/search/?api=1&query=" + quote_plus(q)
                coord, query = "", ""

            stats[lvl] += 1
            fill = PatternFill("solid", fgColor=LEVEL_COLORS[lvl])
            c = ws.cell(r, out_idx["Xarita link"], "Xaritada ochish")
            c.hyperlink = link
            c.font = Font(color="0563C1", underline="single")
            ws.cell(r, out_idx["Koordinata"], coord)
            ws.cell(r, out_idx["Aniqlik"], LEVEL_NAMES[lvl]).fill = fill
            ws.cell(r, out_idx["Topilgan so'rov"], query)
            ws.cell(r, out_idx["Manba"], a.provider if best else "")

            if i % 10 == 0 or i == total:
                el = time.time() - t0
                eta = el / i * (total - i)
                print(f"  {i}/{total}  ({i*100//total}%)  qolgan vaqt ~{int(eta//60)} daq")
            if i % SAVE_EVERY == 0:
                save_cache(cache)
                wb.save(out)
    except KeyboardInterrupt:
        print("\nTo'xtatildi — hozirgacha bo'lganlar saqlanmoqda...")
    finally:
        save_cache(cache)
        wb.save(out)

    print(f"\nSaqlandi: {out}")
    for lvl in sorted(LEVEL_NAMES, reverse=True):
        print(f"  {LEVEL_NAMES[lvl]:<14} {stats[lvl]}")


if __name__ == "__main__":
    main()
