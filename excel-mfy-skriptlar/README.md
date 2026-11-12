# Excel MFY skriptlari

Biznes ro'yxatga olish (YaTT / yuridik shaxslar) Excel fayllari bilan ishlash uchun Python skriptlari:
mahalla nomlarini tozalash, MFY ro'yxati bilan moslash, manzilni bo'lish, F.I.Sh. solishtirish,
fayllarni tuman/mahalla kesimida bo'lish.

## O'rnatish (bir marta)

```
py -m pip install -r requirements.txt
```

## Ishlatish

1. Kerakli skriptni Excel fayl turgan papkaga ko'chiring.
2. Shu papkada: `py skript_nomi.py`
3. Natija yangi faylga/papkaga yoziladi — **asl fayl o'zgarmaydi**.
4. Ishga tushirishdan oldin natija fayli Excel'da **yopiq** bo'lsin.

## Skriptlar

| № | Skript | Nima qiladi | Natija |
|---|---|---|---|
| 1 | `xarita_link.py` | Manzil bo'yicha Google Maps linki (geokodlash) | `<nom>_xarita.xlsx` |
| 2 | `mahalla_nomi.py` | Mahalla nomini `"NOM" MFY` stiliga keltiradi | `<nom>_mahalla.xlsx` |
| 3 | `mfy_moslash.py` | Mahallani **МФЙ лар руйхати** bilan tuman bo'yicha solishtiradi, STIR qo'yadi | `<nom>_mfy.xlsx` |
| 4 | `yuridik_mfy.py` | Yuridik fayliga SOATO + nom bo'yicha MFY nomi va STIR (boshqa tuman mahallasi tushmaydi) | `<nom>_mfy.xlsx` |
| 5 | `yatt_manzil.py` | YaTT reestridan faollarni ajratadi, Яшаш манзилини Viloyat / Tuman / QFY / MFY / Ko'cha / Uy ga bo'ladi | `<nom>_faol.xlsx` |
| 6 | `fish_moslash.py` | Soliq (kirill) F.I.Sh.ni lotinga o'girib, adliya bilan **aniq** moslaydi | `<nom>_natija.xlsx` |
| 7 | `mfy_bolish.py` | Papkadagi barcha fayllarni mahalla bo'yicha papkalarga bo'ladi | `MFY_kesimida\` + `Svod.xlsx` |
| 8 | `tuman_mfy_bolish.py` | Faylni tuman papkasi → `Mahalla_mfy_23.xlsx` (23 = YaTT soni) qilib bo'ladi | `Tuman_MFY_kesimida\` + `Svod.xlsx` |

## Fayl nomi / joyi boshqa bo'lsa

Har bir skript boshida **SOZLAMALAR** bo'limi bor — faqat shu yer o'zgartiriladi.

| Skript | Fayl nomi | Ustun sarlavhalari | Boshqa |
|---|---|---|---|
| `xarita_link.py` | `DEFAULT_INPUT` | `HEADER_NAMES` | `DEFAULT_PROVIDER`, `DEFAULT_KEY`, `DEFAULT_LIMIT`, `BBOX` (viloyat chegarasi) |
| `mahalla_nomi.py` | `DEFAULT_INPUT` | `SOURCE_HEADER`, `NEW_HEADER` | `APOSTROPHE`, `UPPERCASE` |
| `mfy_moslash.py` | `MAIN_FILE`, `REF_FILE` | `MAIN_*_HEADERS`, `REF_*_HEADERS` | `MIN_FILL`, `GREEN_FROM`, `YELLOW_FROM` |
| `yuridik_mfy.py` | `MAIN_FILE`, `REF_FILE` | `H_SOATO`, `H_MFY_NAME`, `H_TUMAN`, `H_MANZIL`, `R_*` | `MIN_NAME` |
| `yatt_manzil.py` | `INPUT_FILE` | `ADDRESS_HEADER`, `STATUS_HEADER` | `ONLY_ACTIVE`, `TRANSLIT` |
| `fish_moslash.py` | `INPUT_FILE` | `SOLIQ_SHEET`, `ADLIYA_SHEET`, `NAME_HEADERS` | `MIN_ACCEPT` (100 = faqat aniq) |
| `mfy_bolish.py` | `FILES` (bo'sh = hammasi) | `MFY_HEADERS` | `OUT_DIR` |
| `tuman_mfy_bolish.py` | `INPUT_FILE` | `TUMAN_HEADERS`, `MFY_HEADERS` | `OUT_DIR` |

**Qoidalar:**

- Fayl topilmasa, ko'p skriptlar papkadagi Excel fayllar ro'yxatini chiqarib, raqam bilan tanlatadi.
- Fayl nomini buyruqda ham berish mumkin: `py fish_moslash.py "Boshqa fayl.xlsx"`
  (`xarita_link.py` da: `py xarita_link.py -i "fayl.xlsx"`).
- Fayl boshqa papkada bo'lsa — to'liq yo'l, oldida **`r`** bilan:
  `INPUT_FILE = r"Z:\МОЛИЯ БУЛИМИ\Samandar\8\Книга6.xlsx"`
- Ustun sarlavhasi boshqacha bo'lsa — `..._HEADERS` ro'yxatiga **kichik harflarda** qo'shing.
