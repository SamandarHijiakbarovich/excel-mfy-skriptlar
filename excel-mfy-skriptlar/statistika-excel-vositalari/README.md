# Statistika Excel vositalari

Surxondaryo viloyat statistika boshqarmasi uchun tayyorlangan Python
skriptlari to'plami. Davlat statistikasi Excel fayllarini qayta ishlash uchun:
MFY (mahalla) nomlarini ajratish, STIR/INN/PINFL bo'yicha moslashtirish,
ustunlarni tozalash, fayllarni tuman va mahalla kesimida bo'lish hamda svodka
(hisobot) jadvallari tayyorlash.

Har bir skript mustaqil ishlaydi — kerakligini olib ishlataverasiz.

---

## O'rnatish (bir marta)

1. [Python 3.9+](https://www.python.org/downloads/) ni o'rnating
   (o'rnatishda **"Add Python to PATH"** ni belgilang).
2. Kerakli kutubxonalarni o'rnating:

   ```bash
   pip install -r requirements.txt
   ```

   yoki alohida:

   ```bash
   pip install openpyxl pandas
   ```

## Qanday ishlatish

Har bir skriptni Excel fayl bilan bitta papkaga qo'yib, terminalda ishga
tushiring. Ikki usul bor:

```bash
# 1-usul: fayl nomini argument qilib berasiz
python 01_bolish/tuman_kesimida_bolish.py "mening_faylim.xlsx"

# 2-usul: skript ichidagi INPUT_FILE sozlamasiga fayl nomini yozib qo'yasiz
python 01_bolish/tuman_kesimida_bolish.py
```

> **Eslatma:** har bir skript boshida `# SOZLAMALAR` bo'limi bor — fayl nomi,
> ustunlar, chiqish papkasi va boshqalarni o'sha yerdan o'zgartirasiz.
> Ishlatishdan oldin kiritiladigan Excel faylni **yopib qo'ying** (Excelda
> ochiq tursa, dastur faylni o'qiy olmaydi).

---

## Skriptlar ro'yxati

### 📂 01_bolish/ — Fayllarni bo'lish

| Skript | Vazifasi |
|--------|----------|
| `tuman_boyicha_bolish.py` | Bitta faylni **tuman** bo'yicha alohida fayllarga bo'ladi. |
| `tuman_kesimida_bolish.py` | Ikkala varaqni (masalan "Сўровномалар" + "Интервьюерлар") tuman kesimida bo'ladi — har tumanga bitta fayl, ichida ikki varaq, oxirida soni. |
| `mahalla_kesimida_bolish.py` | Bir yoki bir nechta faylni `<fayl>/<Tuman>/<Mahalla>.xlsx` ko'rinishida bo'ladi. Har safar chiqish papkasini tozalab qayta yozadi. Keraksiz ustunlarni (`DROP_COLS`) tashlab ketadi. |
| `surxondaryo_tuman_bolish.py` | 2 varaqli (YaTT/Yuridik) faylni tuman bo'yicha bo'ladi; tuman kodidan nomga o'tkazadi. |
| `yatt_mahalla_bolish.py` | YaTT (jismoniy) faylni `YaTT/Tumanlar/<Tuman>/<Mahalla>_Qo'shimcha_yatt.xlsx` tuzilmasida bo'ladi + `00_Jami_royxat.xlsx`. |
| `yuridik_mahalla_bolish.py` | Yuridik faylni `Yuridik/Tumanlar/<Tuman>/<Mahalla>_Qo'shimcha_yuridik.xlsx` tuzilmasida bo'ladi + jami ro'yxat. |

### 📂 02_moslashtirish/ — Yozuvlarni ulash (STIR/INN/PINFL/manzil)

| Skript | Vazifasi |
|--------|----------|
| `stir_moslash.py` | Faylga F.I.Sh (ism) bo'yicha STIR olib qo'yadi. O'zbekcha ism moslashtirish: kiril↔lotin, o'g'li/qizi, jins, fuzzy + manzil bo'yicha aniqlashtirish. Shubhali mosliklarni "Tekshirish" varag'iga chiqaradi. |
| `stir_gaz.py` | `stir_moslash` ning gaz fayli uchun sozlangan varianti. |
| `stir_suv.py` | `stir_moslash` ning suv fayli uchun sozlangan varianti. |
| `stir_qoshish.py` | 1722-manba fayldan MFY/tashkilot STIR raqamini olib qo'shadi. |
| `inn_moslash.py` | Korxona (yuridik) nomi bo'yicha INN moslashtirib qo'yadi. |
| `pinfl_olib_otish.py` | F.I.Sh bo'yicha PINFL (JSHSHIR) ni bir fayldan ikkinchisiga olib o'tadi. |
| `yatt_manzil_telefon.py` | F.I.Sh bo'yicha manzil va telefon raqamini olib o'tadi. |

### 📂 03_tozalash/ — Ustun va qatorlarni tozalash

| Skript | Vazifasi |
|--------|----------|
| `mfy_nomini_ajratish.py` | To'liq manzildan MFY (mahalla) nomini ajratib alohida ustunga chiqaradi. |
| `yuridik_shakl_olib_tashlash.py` | Tashkilot nomidan "MChJ", "XK", "OK" va shu kabi tashkiliy-huquqiy shakllarni olib tashlaydi (tez versiya). |
| `mchj_xk_ok_qator_ochirish.py` | Nomida MChJ/XK/OK bo'lgan **qatorlarni** o'chiradi (tez versiya). |

### 📂 04_svodka/ — Hisobot jadvallari

| Skript | Vazifasi |
|--------|----------|
| `svodka_yatt_yuridik.py` | Tumanlar kesimida YaTT (jismoniy) va yuridik shaxslar sonini bitta chiroyli jadvalga chiqaradi. Faylni bo'lmaydi, faqat svodka. |
| `xatolar_svodka.py` | Yuklab olingan "xatolar" fayli asosida har tuman uchun formatlangan svodka tayyorlaydi (sarlavha, sana, yashil shapka, JAMI qatori, rangli "Tuzatilgan ulushi" ustuni) + `00_Jami_svodka.xlsx`. |

### 📂 05_boshqa/

| Skript | Vazifasi |
|--------|----------|
| `hash_ochish.py` | Hashlangan faylni **sizdagi kalit/jadval** bilan ulaydi (REGIM 1 — hash→qiymat jadvali, REGIM 2 — o'z bazangizni SHA-256 bilan qayta hashlab ulash). Bu hashni "sindirish" emas — faqat sizda bor kalit bilan moslash. |

### 📂 ishga_tushirish/

| Fayl | Vazifasi |
|------|----------|
| `Svodka_ishga_tushirish.bat` | Python bilmagan hamkasblar uchun — ikki marta bosib svodka skriptini ishga tushiradigan Windows yorliq fayli. |

---

## Umumiy xususiyatlar

- **Kiril↔Lotin** transliteratsiya bilan ism/nom moslashtirish.
- Katta fayllar uchun `openpyxl` `read_only`/`write_only` rejimlari.
- Chiroyli formatlash: rangli shapka, chegaralar, `JAMI` qatori, avto-filtr,
  `freeze_panes`, shartli rang (ColorScaleRule).
- Har bir skript xatolarni ushlaydi va tushunarli xabar beradi (masalan, fayl
  Excelda ochiq bo'lsa ogohlantiradi).

## Talablar

- Python 3.9 yoki undan yuqori
- `openpyxl`, `pandas` (qarang: `requirements.txt`)
