# -*- coding: utf-8 -*-
"""
SOTUV BOT — Telegram bot orqali kunlik sotuv va xarajatlarni Excel faylga yozib boradi.

O'RNATISH:
    pip install python-telegram-bot==21.4 openpyxl

FAQAT TANLANGAN ODAMLAR ISHLATISHI UCHUN:
    ALLOWED_USERS nomli environment variable orqali ruxsat berilgan
    Telegram ID'larni vergul bilan yozing, masalan: "111111111,222222222"
    (o'z ID'ingizni bilish uchun Telegram'da @userinfobot ga /start yozing).
    Bo'sh qoldirilsa — bot hammaga ochiq bo'ladi.

ISHGA TUSHIRISH:
    1. Pastdagi BOT_TOKEN o'rniga o'z bot tokeningizni yozing (BotFather'dan olinadi)
    2. EXCEL_FILE nomini xohlagancha o'zgartiring
    3. python sotuv_bot.py

BOT BUYRUQLARI:
    /start      - botni boshlash, menyu ko'rsatish
    /sotuv      - yangi sotuvni bosqichma-bosqich yozish (mahsulot -> miqdor -> narx)
    /xarajat    - yangi xarajatni bosqichma-bosqich yozish (nomi -> summa)
    /qoldik     - barcha mahsulotlarning joriy qoldig'ini ko'rsatish
    /mahsulot   - yangi mahsulot qo'shish (nomi -> boshlang'ich miqdor -> narx)
    /bekor      - joriy amalni bekor qilish

ERKIN MATN (bitta xabar bilan yozish, komandasiz):
    Shunchaki bitta xabar yuborsangiz ham bot tushunadi:
        "loro gold khaki sotildi 1 ta narxi 380 ming"
        "krem futболка sotildi 3 dona narxi 45000"
        "xarajat benzin uchun 50 ming"
    "ming" so'zi 1000 ga, "mln"/"million" so'zi 1 000 000 ga ko'paytiriladi.
    Agar bot matnni tushunmasa, /sotuv yoki /xarajat buyrug'idan bosqichma-bosqich
    kiritishni taklif qiladi.

Excel fayl 3 varaqdan iborat bo'ladi:
    - Mahsulotlar : Nomi | Boshlang'ich | Sotilgan | Qoldiq | Narx
    - Sotuvlar    : Sana | Xodim | Mahsulot | Miqdor | Narx | Summa
    - Xarajatlar  : Sana | Xodim | Nomi | Summa
"""

import os
import re
import datetime
from zoneinfo import ZoneInfo
from html import escape
from openpyxl import Workbook, load_workbook

from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    Update,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ============ SOZLAMALAR ============
# Token kodning ichiga yozilmaydi — xavfsizlik uchun tashqi "environment variable"dan olinadi.
# Kompyuterda sinab ko'rish uchun pastdagi qatorni vaqtincha o'zgartirib turishingiz ham mumkin:
#   BOT_TOKEN = "123456:ABC-tokeningiz"
BOT_TOKEN = os.environ.get("BOT_TOKEN", "BU_YERGA_TOKENINGIZNI_YOZING")
EXCEL_FILE = "sotuv_hisobot.xlsx"

# Botdan faqat shu Telegram ID'ga ega odamlar foydalana oladi.
# ID'larni vergul bilan ajratib yozing, masalan: "111111111,222222222"
# Bo'sh qoldirsangiz — hamma foydalana oladi (tavsiya etilmaydi).
# O'z ID'ingizni bilish uchun Telegram'da @userinfobot ga /start yozing.
_allowed_raw = os.environ.get("ALLOWED_USERS", "")
ALLOWED_USERS = {int(x) for x in _allowed_raw.split(",") if x.strip().isdigit()}


def is_allowed(update: Update) -> bool:
    if not ALLOWED_USERS:
        return True
    return update.effective_user.id in ALLOWED_USERS

# ConversationHandler holatlari
MODEL_STEP, PRODUCT, QTY, PRICE = range(4)
EXP_NAME, EXP_SUM = range(4, 6)
NEW_NAME, NEW_QTY, NEW_PRICE = range(6, 9)


# ============ EXCEL YORDAMCHI FUNKSIYALAR ============
def ensure_excel():
    """Agar fayl mavjud bo'lmasa, kerakli varaqlar bilan yaratadi."""
    if os.path.exists(EXCEL_FILE):
        return
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Mahsulotlar"
    ws1.append(["Nomi", "Boshlang'ich", "Sotilgan", "Qoldiq", "Narx"])

    ws2 = wb.create_sheet("Sotuvlar")
    ws2.append(["Sana", "Xodim", "Mahsulot", "Miqdor", "Narx", "Summa"])

    ws3 = wb.create_sheet("Xarajatlar")
    ws3.append(["Sana", "Xodim", "Nomi", "Summa"])

    wb.save(EXCEL_FILE)


def find_product_row(ws, name):
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=1).value
        if cell and str(cell).strip().lower() == name.strip().lower():
            return row
    return None


def add_product(name, boshlangich, narx):
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Mahsulotlar"]
    row = find_product_row(ws, name)
    if row:
        ws.cell(row=row, column=2, value=boshlangich)
        ws.cell(row=row, column=5, value=narx)
        ws.cell(row=row, column=4, value=boshlangich - (ws.cell(row=row, column=3).value or 0))
    else:
        new_row = ws.max_row + 1
        ws.cell(row=new_row, column=1, value=name)
        ws.cell(row=new_row, column=2, value=boshlangich)
        ws.cell(row=new_row, column=3, value=0)
        ws.cell(row=new_row, column=4, value=boshlangich)
        ws.cell(row=new_row, column=5, value=narx)
    wb.save(EXCEL_FILE)


def record_sale(xodim, mahsulot, miqdor, narx):
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws_m = wb["Mahsulotlar"]
    ws_s = wb["Sotuvlar"]

    row = find_product_row(ws_m, mahsulot)
    if row is None:
        # mahsulot ro'yxatda yo'q bo'lsa, 0 boshlang'ich bilan qo'shamiz
        row = ws_m.max_row + 1
        ws_m.cell(row=row, column=1, value=mahsulot)
        ws_m.cell(row=row, column=2, value=0)
        ws_m.cell(row=row, column=3, value=0)
        ws_m.cell(row=row, column=4, value=0)
        ws_m.cell(row=row, column=5, value=narx)

    sotilgan = (ws_m.cell(row=row, column=3).value or 0) + miqdor
    boshlangich = ws_m.cell(row=row, column=2).value or 0
    ws_m.cell(row=row, column=3, value=sotilgan)
    ws_m.cell(row=row, column=4, value=boshlangich - sotilgan)
    if narx:
        ws_m.cell(row=row, column=5, value=narx)

    summa = miqdor * narx
    ws_s.append([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), xodim, mahsulot, miqdor, narx, summa])
    wb.save(EXCEL_FILE)
    return boshlangich - sotilgan


def record_expense(xodim, nomi, summa):
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Xarajatlar"]
    ws.append([datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), xodim, nomi, summa])
    wb.save(EXCEL_FILE)


def get_product_names():
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Mahsulotlar"]
    names = []
    for row in range(2, ws.max_row + 1):
        nomi = ws.cell(row=row, column=1).value
        if nomi:
            names.append(str(nomi))
    return names


def _parse_sana(sana_str):
    try:
        return datetime.datetime.strptime(sana_str, "%Y-%m-%d %H:%M")
    except (ValueError, TypeError):
        return datetime.datetime.min


def undo_last_action():
    """Eng oxirgi yozilgan sotuv yoki xarajatni bekor qiladi (o'chiradi).
    Sotuv bo'lsa — mahsulot qoldig'ini ham qaytaradi.
    Natija: tavsif matni (string) yoki None (bekor qilinadigan narsa topilmasa)."""
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws_s = wb["Sotuvlar"]
    ws_x = wb["Xarajatlar"]
    ws_m = wb["Mahsulotlar"]

    last_sale = None
    if ws_s.max_row > 1:
        row = ws_s.max_row
        last_sale = (row, [ws_s.cell(row=row, column=c).value for c in range(1, 7)])

    last_exp = None
    if ws_x.max_row > 1:
        row = ws_x.max_row
        last_exp = (row, [ws_x.cell(row=row, column=c).value for c in range(1, 5)])

    if not last_sale and not last_exp:
        return None

    sale_time = _parse_sana(last_sale[1][0]) if last_sale else datetime.datetime.min
    exp_time = _parse_sana(last_exp[1][0]) if last_exp else datetime.datetime.min

    if sale_time >= exp_time:
        row, (sana, xodim, mahsulot, miqdor, narx, summa) = last_sale
        ws_s.delete_rows(row)
        prow = find_product_row(ws_m, mahsulot)
        if prow:
            sotilgan = (ws_m.cell(row=prow, column=3).value or 0) - miqdor
            boshlangich = ws_m.cell(row=prow, column=2).value or 0
            ws_m.cell(row=prow, column=3, value=sotilgan)
            ws_m.cell(row=prow, column=4, value=boshlangich - sotilgan)
        wb.save(EXCEL_FILE)
        return f"🗑 Bekor qilindi: {mahsulot} — {miqdor} dona x {narx:,.0f} so'm ({xodim})"
    else:
        row, (sana, xodim, nomi, summa) = last_exp
        ws_x.delete_rows(row)
        wb.save(EXCEL_FILE)
        return f"🗑 Bekor qilindi: xarajat — {nomi} — {summa:,.0f} so'm ({xodim})"


def get_report_for_date(date_str):
    """Berilgan sana (YYYY-MM-DD) uchun sotuvlar va xarajatlar hisobotini tayyorlaydi."""
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws_s = wb["Sotuvlar"]
    ws_x = wb["Xarajatlar"]

    sotuvlar = []
    for row in range(2, ws_s.max_row + 1):
        sana = ws_s.cell(row=row, column=1).value
        if sana and str(sana).startswith(date_str):
            sotuvlar.append([ws_s.cell(row=row, column=c).value for c in range(1, 7)])

    xarajatlar = []
    for row in range(2, ws_x.max_row + 1):
        sana = ws_x.cell(row=row, column=1).value
        if sana and str(sana).startswith(date_str):
            xarajatlar.append([ws_x.cell(row=row, column=c).value for c in range(1, 5)])

    lines = [f"📅 <b>{date_str}</b> uchun hisobot:\n"]

    if sotuvlar:
        jami_dona = sum(s[3] for s in sotuvlar)
        jami_summa = sum(s[5] for s in sotuvlar)
        lines.append(f"🛒 <b>Sotuvlar</b> — {len(sotuvlar)} ta yozuv, jami {jami_dona:,.0f} dona, {jami_summa:,.0f} so'm:")
        for sana, xodim, mahsulot, miqdor, narx, summa in sotuvlar:
            vaqt = str(sana).split(" ")[-1] if " " in str(sana) else ""
            lines.append(f"  {vaqt}  {mahsulot} — {miqdor} dona x {narx:,.0f} = {summa:,.0f} so'm ({xodim})")
    else:
        lines.append("🛒 Sotuvlar: bu kunda yozuv yo'q")

    lines.append("")

    if xarajatlar:
        jami_x = sum(x[3] for x in xarajatlar)
        lines.append(f"💸 <b>Xarajatlar</b> — {len(xarajatlar)} ta yozuv, jami {jami_x:,.0f} so'm:")
        for sana, xodim, nomi, summa in xarajatlar:
            vaqt = str(sana).split(" ")[-1] if " " in str(sana) else ""
            lines.append(f"  {vaqt}  {nomi} — {summa:,.0f} so'm ({xodim})")
    else:
        lines.append("💸 Xarajatlar: bu kunda yozuv yo'q")

    text = "\n".join(lines)
    return text


def get_report_date_markup():
    """Oxirgi 7 kun uchun tugmalar + 'boshqa sana' tugmasi."""
    today = datetime.date.today()
    buttons = []
    for i in range(7):
        d = today - datetime.timedelta(days=i)
        label = "Bugun" if i == 0 else ("Kecha" if i == 1 else d.strftime("%d.%m"))
        buttons.append([InlineKeyboardButton(label, callback_data=f"rep:{d.isoformat()}")])
    buttons.append([InlineKeyboardButton("✍️ Boshqa sana yozish", callback_data="rep:custom")])
    return InlineKeyboardMarkup(buttons)


DATE_PATTERNS = [
    (r"^(\d{4})-(\d{1,2})-(\d{1,2})$", lambda m: (int(m[1]), int(m[2]), int(m[3]))),
    (r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", lambda m: (int(m[3]), int(m[2]), int(m[1]))),
    (r"^(\d{1,2})\.(\d{1,2})$", lambda m: (datetime.date.today().year, int(m[2]), int(m[1]))),
    (r"^(\d{1,2})/(\d{1,2})/(\d{4})$", lambda m: (int(m[3]), int(m[2]), int(m[1]))),
]


def parse_date_text(text):
    """'07.09', '07.09.2026', '2026-09-07' kabi matnlarni sanaga aylantiradi.
    Muvaffaqiyatsiz bo'lsa None qaytaradi."""
    t = text.strip()
    for pattern, extractor in DATE_PATTERNS:
        m = re.match(pattern, t)
        if m:
            try:
                year, month, day = extractor(m)
                return datetime.date(year, month, day)
            except ValueError:
                return None
    return None


def group_products_by_model():
    """Mahsulotlar varag'idagi har bir qatorni tegishli modelga ajratadi.
    Natija: {model: [(rang, qoldiq, boshlangich), ...]}"""
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Mahsulotlar"]

    # Eng uzun nomdan boshlab tekshiramiz, shunda "LP 19 Cеребреный с точкой"
    # "LP 19 Cеребреный" bilan aralashib ketmaydi.
    models_by_len = sorted(MODEL_PREFIXES, key=len, reverse=True)

    groups = {}
    for row in range(2, ws.max_row + 1):
        nomi = ws.cell(row=row, column=1).value
        if not nomi:
            continue
        nomi = str(nomi)
        boshlangich = ws.cell(row=row, column=2).value or 0
        qoldiq = ws.cell(row=row, column=4).value or 0

        model = next((m for m in models_by_len if nomi.startswith(m)), None)
        if model:
            rang = nomi[len(model):].strip() or "-"
        else:
            model, rang = nomi, "-"

        groups.setdefault(model, []).append((rang, qoldiq, boshlangich))
    return groups


def format_model_block(model, items):
    width = max((len(r) for r, _, _ in items), default=4)
    lines = [f"{model}"]
    for rang, qoldiq, boshlangich in items:
        lines.append(f"  {rang.ljust(width)}  {qoldiq}/{boshlangich}")
    return "\n".join(lines)


def get_qoldiq_models_markup():
    """Har bir model uchun alohida tugma bilan inline klaviatura qaytaradi."""
    groups = group_products_by_model()
    order = [m for m in MODEL_DISPLAY_ORDER if m in groups] + [
        m for m in groups if m not in MODEL_DISPLAY_ORDER
    ]
    buttons = [[InlineKeyboardButton(model, callback_data=f"qm:{model}")] for model in order]
    return InlineKeyboardMarkup(buttons), groups


# Nechta va undan kam qolganda "kam qolgan" hisoblanadi. Xohlasangiz
# Railway'da LOW_STOCK_THRESHOLD nomli environment variable orqali o'zgartirish mumkin.
LOW_STOCK_THRESHOLD = int(os.environ.get("LOW_STOCK_THRESHOLD", "2"))


def get_low_stock_text():
    """Qoldig'i LOW_STOCK_THRESHOLD dan kam yoki teng bo'lgan mahsulotlarni,
    model bo'yicha guruhlab qaytaradi."""
    groups = group_products_by_model()
    order = [m for m in MODEL_DISPLAY_ORDER if m in groups] + [
        m for m in groups if m not in MODEL_DISPLAY_ORDER
    ]

    blocks = []
    for model in order:
        kam = [(r, q, b) for r, q, b in groups[model] if q <= LOW_STOCK_THRESHOLD]
        if kam:
            blocks.append(format_model_block(model, kam))

    if not blocks:
        return f"✅ Hozircha {LOW_STOCK_THRESHOLD} tadan kam qolgan mahsulot yo'q."

    text = "\n\n".join(blocks)
    return (
        f"⚠️ <b>Kam qolgan mahsulotlar</b> ({LOW_STOCK_THRESHOLD} tadan kam, qoldiq/boshlang'ich):\n"
        + "<pre>" + escape(text) + "</pre>"
    )





# ============ ERKIN MATNNI TUSHUNISH (parser) ============
UNIT_MULT = {
    "ming": 1_000,
    "mln": 1_000_000,
    "million": 1_000_000,
    "mln.": 1_000_000,
}

QTY_WORDS = r"(?:ta|dona)"
SOLD_WORDS = r"(?:sotildi|sotdim|sotildim|sotib\s+yubordim)"
PRICE_WORDS = r"(?:narxi|narxida|narx|baho|bahosi)"
EXPENSE_WORDS = r"(?:xarajat|xarajatga|harajat|rasxod)"


def _parse_amount(num_str, unit_str):
    num_str = num_str.replace(",", ".")
    val = float(num_str)
    if unit_str:
        val *= UNIT_MULT.get(unit_str.lower().strip("."), 1)
    return val


def parse_sale_text(text: str):
    """'loro gold khaki sotildi 1 ta narxi 380 ming' kabi xabarni ajratadi.
    Muvaffaqiyatli bo'lsa (mahsulot, miqdor, narx) qaytaradi, aks holda None."""
    t = text.strip()
    sold_match = re.search(SOLD_WORDS, t, re.IGNORECASE)
    if not sold_match:
        return None

    mahsulot = t[: sold_match.start()].strip(" .,:-")
    rest = t[sold_match.end():]
    if not mahsulot:
        return None

    qty_match = re.search(r"(\d+(?:[.,]\d+)?)\s*" + QTY_WORDS, rest, re.IGNORECASE)
    price_match = re.search(
        PRICE_WORDS + r"\D{0,5}(\d+(?:[.,]\d+)?)\s*(ming|mln\.?|million)?",
        rest,
        re.IGNORECASE,
    )
    if not qty_match or not price_match:
        return None

    miqdor = _parse_amount(qty_match.group(1), None)
    narx = _parse_amount(price_match.group(1), price_match.group(2))
    return mahsulot, miqdor, narx


def parse_expense_text(text: str):
    """'xarajat benzin uchun 50 ming' kabi xabarni ajratadi.
    Muvaffaqiyatli bo'lsa (nomi, summa) qaytaradi, aks holda None."""
    t = text.strip()
    exp_match = re.search(EXPENSE_WORDS, t, re.IGNORECASE)
    if not exp_match:
        return None

    amount_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(ming|mln\.?|million)?\s*$", t, re.IGNORECASE)
    if not amount_match:
        # summa oxirida bo'lmasligi mumkin, butun matndan qidiramiz
        amount_match = re.search(r"(\d+(?:[.,]\d+)?)\s*(ming|mln\.?|million)?", t, re.IGNORECASE)
    if not amount_match:
        return None

    summa = _parse_amount(amount_match.group(1), amount_match.group(2))

    nomi = t[: exp_match.start()] + t[exp_match.end(): amount_match.start()]
    nomi = re.sub(r"\b(uchun|ga|uchn)\b", " ", nomi, flags=re.IGNORECASE)
    nomi = nomi.strip(" .,:-")
    if not nomi:
        nomi = "Xarajat"
    return nomi, summa


async def handle_free_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Foydalanuvchi komandasiz, bitta xabar bilan sotuv/xarajat yozganda ishlaydi."""
    if not is_allowed(update):
        return  # ruxsatsiz odamga hatto javob ham bermaymiz

    if await hisobot_by_text(update, context):
        return

    text = update.message.text
    xodim = xodim_ismi(update)

    exp = parse_expense_text(text)
    if exp and re.search(EXPENSE_WORDS, text, re.IGNORECASE):
        nomi, summa = exp
        record_expense(xodim, nomi, summa)
        await update.message.reply_text(f"✅ Xarajat yozildi: {nomi} — {summa:,.0f} so'm")
        return

    sale = parse_sale_text(text)
    if sale:
        mahsulot, miqdor, narx = sale
        qoldiq = record_sale(xodim, mahsulot, miqdor, narx)
        ogohlantirish = (
            f"\n\n⚠️ Diqqat! Faqat {qoldiq:,.0f} dona qoldi."
            if qoldiq <= LOW_STOCK_THRESHOLD else ""
        )
        await update.message.reply_text(
            f"✅ Yozildi: {mahsulot} — {miqdor:,.0f} dona x {narx:,.0f} so'm\n"
            f"Qolgan qoldiq: {qoldiq:,.0f} dona{ogohlantirish}"
        )
        return

    await update.message.reply_text(
        "Tushunmadim 🙁 Quyidagicha yozib ko'ring:\n"
        "  \"mahsulot nomi sotildi 1 ta narxi 380 ming\"\n"
        "  \"xarajat benzin uchun 50 ming\"\n"
        "Yoki /sotuv, /xarajat buyrug'idan foydalaning."
    )




# ============ BOSHLANG'ICH MAHSULOTLAR (bir martalik seed) ============
# Bu ro'yxat foydalanuvchining eski Excel hisobotidan olingan.
# Format: (Nomi, Boshlang'ich, Sotilgan, Qoldiq, Narx)
# Mahsulot nomlaridagi asosiy modellar ro'yxati — /qoldik'da guruhlash uchun ishlatiladi.
MODEL_PREFIXES = [
    "Сумка LP 19 золтой",
    "Сумка LP 19 Cеребреный с точкой",
    "Сумка LP 19 Cеребреный",
    "Сумка LP 19 Золотой с точкой",
    "Сумка ALEX MIA CD-9358",
    "Сумка ALEX MIA CD-9838",
    "Сумка ALEX MIA CD-9681",
    "Сумка Balenciaga",
    "Сумка LV Man",
]


# /qoldik tugmalarida ko'rsatish tartibi (foydalanuvchi belgilagan ketma-ketlik).
MODEL_DISPLAY_ORDER = [
    "Сумка LP 19 золтой",
    "Сумка LP 19 Cеребреный",
    "Сумка LP 19 Золотой с точкой",
    "Сумка LP 19 Cеребреный с точкой",
    "Сумка ALEX MIA CD-9358",
    "Сумка ALEX MIA CD-9838",
    "Сумка ALEX MIA CD-9681",
    "Сумка Balenciaga",
    "Сумка LV Man",
]

SEED_PRODUCTS = [
    ("Сумка LP 19 золтой Шоколадный", 5, 1, 4, 379000),
    ("Сумка LP 19 золтой Черный", 5, 2, 3, 379000),
    ("Сумка LP 19 золтой Кремовый", 5, 2, 3, 379000),
    ("Сумка LP 19 золтой Песочно-серый", 5, 3, 2, 379000),
    ("Сумка LP 19 золтой Светло-синий", 5, 0, 5, 379000),
    ("Сумка LP 19 золтой Электро-синий", 3, 0, 3, 379000),
    ("Сумка LP 19 золтой Ночной-синий", 5, 0, 5, 379000),
    ("Сумка LP 19 золтой Светло-зелёный", 5, 0, 5, 379000),
    ("Сумка LP 19 золтой Розовый", 3, 1, 2, 379000),
    ("Сумка LP 19 золтой Золотая-пальма", 3, 0, 3, 379000),
    ("Сумка LP 19 золтой Слоновый пепел", 5, 0, 5, 379000),
    ("Сумка LP 19 золтой Красное вино", 5, 2, 3, 379000),
    ("Сумка LP 19 золтой Военно-Зелёный", 5, 3, 2, 379000),
    ("Сумка LP 19 Cеребреный Черный", 5, 1, 4, 378000),
    ("Сумка LP 19 Cеребреный Шоколадный", 5, 1, 4, 378000),
    ("Сумка LP 19 Cеребреный Красное вино", 5, 0, 5, 378000),
    ("Сумка LP 19 Cеребреный Золотая-пальма", 3, 1, 2, 378000),
    ("Сумка LP 19 Cеребреный Кремовый", 5, 1, 4, 378000),
    ("Сумка LP 19 Cеребреный Песочно-серый", 5, 2, 3, 378000),
    ("Сумка LP 19 Cеребреный Розовый", 3, 1, 2, 378000),
    ("Сумка LP 19 Cеребреный Светло-синий", 5, 1, 4, 378000),
    ("Сумка LP 19 Cеребреный Светло-зелёный", 5, 0, 5, 378000),
    ("Сумка LP 19 Cеребреный Серый слоновый", 5, 1, 4, 378000),
    ("Сумка LP 19 Cеребреный Военно-Зелёный", 5, 3, 2, 378000),
    ("Сумка LP 19 Cеребреный Ночной-синий", 5, 1, 4, 378000),
    ("Сумка LP 19 Золотой с точкой Военно-Зелёный", 3, 0, 3, 380000),
    ("Сумка LP 19 Золотой с точкой Песочно-серый", 3, 0, 3, 380000),
    ("Сумка LP 19 Золотой с точкой Электро-синий", 3, 1, 2, 380000),
    ("Сумка LP 19 Золотой с точкой Красное вино", 3, 3, 0, 380000),
    ("Сумка LP 19 Золотой с точкой Шоколадный", 3, 0, 3, 380000),
    ("Сумка LP 19 Золотой с точкой Розовый", 1, 0, 1, 380000),
    ("Сумка LP 19 Золотой с точкой Кремовый", 2, 1, 1, 380000),
    ("Сумка LP 19 Cеребреный с точкой Военно-Зелёный", 3, 2, 1, 380000),
    ("Сумка LP 19 Cеребреный с точкой Песочно-серый", 3, 0, 3, 380000),
    ("Сумка LP 19 Cеребреный с точкой Электро-синий", 3, 1, 2, 380000),
    ("Сумка LP 19 Cеребреный с точкой Красное вино", 3, 0, 3, 380000),
    ("Сумка LP 19 Cеребреный с точкой Шоколадный", 3, 0, 3, 380000),
    ("Сумка LP 19 Cеребреный с точкой Розовый", 1, 1, 0, 380000),
    ("Сумка LP 19 Cеребреный с точкой Кремовый", 2, 1, 1, 380000),
    ("Сумка ALEX MIA CD-9358 Black", 28, 7, 21, 234000),
    ("Сумка ALEX MIA CD-9358 Coffee", 6, 4, 2, 234000),
    ("Сумка ALEX MIA CD-9358 Khkai", 6, 0, 6, 234000),
    ("Сумка ALEX MIA CD-9358 Brown", 3, 3, 0, 234000),
    ("Сумка ALEX MIA CD-9358 Wine", 4, 3, 1, 234000),
    ("Сумка ALEX MIA CD-9358 Blue", 3, 3, 0, 234000),
    ("Сумка ALEX MIA CD-9838 Black", 13, 2, 11, 350000),
    ("Сумка ALEX MIA CD-9838 Coffee", 4, 0, 4, 350000),
    ("Сумка ALEX MIA CD-9838 Brown", 2, 0, 2, 350000),
    ("Сумка ALEX MIA CD-9838 Green", 2, 0, 2, 350000),
    ("Сумка ALEX MIA CD-9838 Mud", 3, 0, 3, 350000),
    ("Сумка ALEX MIA CD-9681 Black", 15, 0, 15, 315000),
    ("Сумка ALEX MIA CD-9681 White", 3, 0, 3, 315000),
    ("Сумка ALEX MIA CD-9681 Beige", 2, 1, 1, 315000),
    ("Сумка ALEX MIA CD-9681 Grey", 3, 0, 3, 315000),
    ("Сумка ALEX MIA CD-9681 Apricot", 1, 1, 0, 315000),
    ("Сумка ALEX MIA CD-9681 Khkai", 8, 0, 8, 315000),
    ("Сумка Balenciaga Черный", 5, 4, 1, 572000),
    ("Сумка Balenciaga Шоколадный", 5, 5, 0, 572000),
    ("Сумка Balenciaga Красное вино", 5, 1, 4, 572000),
    ("Сумка LV Man Black", 15, 0, 15, 300000),
    ("Сумка LV Man Blue", 5, 0, 5, 300000),

]


def seed_products_if_empty():
    """Agar Mahsulotlar varag'i bo'sh bo'lsa (faqat sarlavha qatori bo'lsa),
    eski Excel hisobotidan olingan boshlang'ich ro'yxatni avtomatik yozadi.
    Bu faqat BIR MARTA, fayl yangi yaratilganda ishlaydi — keyingi ishga
    tushirishlarda mahsulotlar allaqachon bor bo'lgani uchun qayta yozilmaydi."""
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Mahsulotlar"]
    if ws.max_row > 1:
        return  # allaqachon mahsulotlar bor, qayta yozmaymiz
    for nomi, boshlangich, sotilgan, qoldiq, narx in SEED_PRODUCTS:
        ws.append([nomi, boshlangich, sotilgan, qoldiq, narx])
    wb.save(EXCEL_FILE)
    print(f"{len(SEED_PRODUCTS)} ta boshlang'ich mahsulot yuklandi.")


def xodim_ismi(update: Update) -> str:
    user = update.effective_user
    return user.full_name or (f"@{user.username}" if user.username else str(user.id))


# ============ ASOSIY MENYU (doim ko'rinib turadigan tugmalar) ============
BTN_SOTUV = "🛒 Sotuv qo'shish"
BTN_XARAJAT = "💸 Xarajat qo'shish"
BTN_QOLDIK = "📦 Qoldiqni ko'rish"
BTN_MAHSULOT = "➕ Yangi mahsulot"
BTN_UNDO = "↩️ Oxirgisini bekor qilish"
BTN_HISOBOT = "📊 Hisobot (sana bo'yicha)"
BTN_KAM_QOLDIQ = "⚠️ Kam qolgan mahsulotlar"

MAIN_MENU_MARKUP = ReplyKeyboardMarkup(
    [
        [BTN_SOTUV, BTN_XARAJAT],
        [BTN_QOLDIK, BTN_MAHSULOT],
        [BTN_HISOBOT, BTN_KAM_QOLDIQ],
        [BTN_UNDO],
    ],
    resize_keyboard=True,
)


# ============ /start ============
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    await update.message.reply_text(
        "Assalomu alaykum! Men sotuv va xarajat hisobchi botiman.\n\n"
        "Pastdagi tugmalardan foydalaning, yoki shunchaki\n"
        "\"mahsulot nomi sotildi 1 ta narxi 380 ming\" deb yozing.",
        reply_markup=MAIN_MENU_MARKUP,
    )


# ============ /sotuv oqimi ============
BTN_QOLDA_YOZISH = "✍️ Qo'lda yozish"


async def sotuv_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return ConversationHandler.END

    context.user_data.pop("model", None)
    names = get_product_names()
    # Faqat mavjud mahsulotlarda ishlatilgan modellarni ko'rsatamiz, tartib
    # MODEL_DISPLAY_ORDER ro'yxatidagi ketma-ketlikda (ya'ni Excel'dagi tartibda).
    used_models = [m for m in MODEL_DISPLAY_ORDER if any(n.startswith(m) for n in names)]

    if used_models:
        keyboard = [[m] for m in used_models] + [[BTN_QOLDA_YOZISH]]
        markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
        await update.message.reply_text("Qaysi model sotildi?", reply_markup=markup)
        return MODEL_STEP

    await update.message.reply_text(
        "Qaysi mahsulot sotildi? (nomini yozing)", reply_markup=ReplyKeyboardRemove()
    )
    return PRODUCT


async def sotuv_model(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    if text == BTN_QOLDA_YOZISH:
        await update.message.reply_text(
            "Mahsulot nomini yozing:", reply_markup=ReplyKeyboardRemove()
        )
        return PRODUCT

    model = next((m for m in MODEL_DISPLAY_ORDER if m == text), None)
    if model is None:
        await update.message.reply_text("Iltimos, ro'yxatdan tanlang.")
        return MODEL_STEP

    context.user_data["model"] = model
    names = get_product_names()
    colors = [n[len(model):].strip() for n in names if n.startswith(model)]
    colors = [c for c in colors if c]  # bo'sh bo'lmaganlari

    if colors:
        keyboard = [[c] for c in colors]
        markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
        await update.message.reply_text("Qaysi rang?", reply_markup=markup)
        return PRODUCT

    await update.message.reply_text(
        "Rangini yozing:", reply_markup=ReplyKeyboardRemove()
    )
    return PRODUCT


async def sotuv_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    model = context.user_data.get("model")
    context.user_data["mahsulot"] = f"{model} {text}" if model else text
    await update.message.reply_text("Nechta dona sotildi?", reply_markup=ReplyKeyboardRemove())
    return QTY


async def sotuv_qty(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["miqdor"] = float(update.message.text.replace(",", "."))
    except ValueError:
        await update.message.reply_text("Iltimos, faqat son kiriting. Nechta dona sotildi?")
        return QTY
    await update.message.reply_text("Bir donasi necha so'mdan sotildi?")
    return PRICE


async def sotuv_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        narx = float(update.message.text.replace(",", "."))
    except ValueError:
        await update.message.reply_text("Iltimos, faqat son kiriting. Narxi qancha?")
        return PRICE

    mahsulot = context.user_data["mahsulot"]
    miqdor = context.user_data["miqdor"]
    xodim = xodim_ismi(update)
    qoldiq = record_sale(xodim, mahsulot, miqdor, narx)
    ogohlantirish = (
        f"\n\n⚠️ Diqqat! Faqat {qoldiq:,.0f} dona qoldi."
        if qoldiq <= LOW_STOCK_THRESHOLD else ""
    )

    await update.message.reply_text(
        f"✅ Yozildi: {mahsulot} — {miqdor} dona x {narx:,.0f} so'm\n"
        f"Qolgan qoldiq: {qoldiq:,.0f} dona{ogohlantirish}",
        reply_markup=MAIN_MENU_MARKUP,
    )
    return ConversationHandler.END


# ============ /xarajat oqimi ============
async def xarajat_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return ConversationHandler.END
    await update.message.reply_text("Xarajat nimaga ketdi? (nomini yozing)")
    return EXP_NAME


async def xarajat_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["exp_nomi"] = update.message.text.strip()
    await update.message.reply_text("Summasi qancha?")
    return EXP_SUM


async def xarajat_sum(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        summa = float(update.message.text.replace(",", "."))
    except ValueError:
        await update.message.reply_text("Iltimos, faqat son kiriting. Summasi qancha?")
        return EXP_SUM
    xodim = xodim_ismi(update)
    record_expense(xodim, context.user_data["exp_nomi"], summa)
    await update.message.reply_text(
        f"✅ Xarajat yozildi: {context.user_data['exp_nomi']} — {summa:,.0f} so'm",
        reply_markup=MAIN_MENU_MARKUP,
    )
    return ConversationHandler.END


# ============ /mahsulot oqimi (yangi mahsulot qo'shish) ============
async def mahsulot_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return ConversationHandler.END
    await update.message.reply_text("Yangi mahsulot nomi?")
    return NEW_NAME


async def mahsulot_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_nomi"] = update.message.text.strip()
    await update.message.reply_text("Boshlang'ich miqdori nechta?")
    return NEW_QTY


async def mahsulot_qty(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        context.user_data["new_qty"] = float(update.message.text.replace(",", "."))
    except ValueError:
        await update.message.reply_text("Iltimos, faqat son kiriting.")
        return NEW_QTY
    await update.message.reply_text("Bir donasi necha so'mdan sotiladi?")
    return NEW_PRICE


async def mahsulot_price(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        narx = float(update.message.text.replace(",", "."))
    except ValueError:
        await update.message.reply_text("Iltimos, faqat son kiriting.")
        return NEW_PRICE
    add_product(context.user_data["new_nomi"], context.user_data["new_qty"], narx)
    await update.message.reply_text(
        f"✅ Mahsulot qo'shildi: {context.user_data['new_nomi']}", reply_markup=MAIN_MENU_MARKUP
    )
    return ConversationHandler.END


# ============ /qoldik ============
async def qoldik(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    markup, groups = get_qoldiq_models_markup()
    if not groups:
        await update.message.reply_text(
            "Hozircha mahsulotlar ro'yxati bo'sh. /mahsulot buyrug'i bilan qo'shing.",
            reply_markup=MAIN_MENU_MARKUP,
        )
        return
    await update.message.reply_text(
        "📦 Qaysi modelning qoldig'ini ko'rmoqchisiz?", reply_markup=markup
    )


async def qoldik_model_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_allowed(update):
        return
    model = query.data[len("qm:"):]
    groups = group_products_by_model()
    items = groups.get(model)
    if not items:
        await query.message.reply_text("Bu model uchun ma'lumot topilmadi.")
        return
    text = format_model_block(model, items)
    await query.message.reply_text("<pre>" + escape(text) + "</pre>", parse_mode="HTML")


# ============ /bekor ============
async def bekor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Bekor qilindi.", reply_markup=MAIN_MENU_MARKUP)
    return ConversationHandler.END


# ============ Oxirgi yozuvni bekor qilish (o'chirish) ============
async def undo_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    natija = undo_last_action()
    if natija is None:
        await update.message.reply_text(
            "Bekor qilinadigan yozuv topilmadi.", reply_markup=MAIN_MENU_MARKUP
        )
    else:
        await update.message.reply_text(natija, reply_markup=MAIN_MENU_MARKUP)


# ============ Hisobot (sana bo'yicha) ============
async def hisobot_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    context.user_data.pop("awaiting_report_date", None)
    await update.message.reply_text(
        "Qaysi sana uchun hisobot kerak?", reply_markup=get_report_date_markup()
    )


async def hisobot_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_allowed(update):
        return

    data = query.data[len("rep:"):]
    if data == "custom":
        context.user_data["awaiting_report_date"] = True
        await query.message.reply_text(
            "Sanani yozing, masalan: 07.09 yoki 07.09.2026"
        )
        return

    await query.message.reply_text(get_report_for_date(data), parse_mode="HTML")


async def hisobot_by_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """context.user_data'da 'awaiting_report_date' bo'lsa, xabarni sana sifatida
    o'qishga urinadi. Agar shu holatda ishlagan bo'lsa True qaytaradi (handle_free_text
    davom etmasligi uchun), aks holda False."""
    if not context.user_data.get("awaiting_report_date"):
        return False

    sana = parse_date_text(update.message.text)
    if sana is None:
        await update.message.reply_text(
            "Sanani tushunmadim. Masalan: 07.09 yoki 07.09.2026 deb yozing."
        )
        return True

    context.user_data.pop("awaiting_report_date", None)
    await update.message.reply_text(
        get_report_for_date(sana.isoformat()), parse_mode="HTML", reply_markup=MAIN_MENU_MARKUP
    )
    return True


# ============ Kam qolgan mahsulotlar ============
async def kam_qoldiq_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    await update.message.reply_text(
        get_low_stock_text(), parse_mode="HTML", reply_markup=MAIN_MENU_MARKUP
    )


# ============ Kunlik avtomatik hisobot ============
async def daily_report_job(context: ContextTypes.DEFAULT_TYPE):
    """Har kuni belgilangan vaqtda barcha ruxsat berilgan foydalanuvchilarga
    o'sha kunning hisobotini avtomatik yuboradi."""
    today = datetime.date.today().isoformat()
    text = get_report_for_date(today)
    kam = get_low_stock_text()
    full_text = text + "\n\n" + kam

    for uid in ALLOWED_USERS:
        try:
            await context.bot.send_message(chat_id=uid, text=full_text, parse_mode="HTML")
        except Exception as e:
            print(f"Kunlik hisobotni {uid}'ga yuborib bo'lmadi: {e}")


def main():
    ensure_excel()
    seed_products_if_empty()
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("qoldik", qoldik))
    app.add_handler(CallbackQueryHandler(qoldik_model_callback, pattern=r"^qm:"))

    sotuv_conv = ConversationHandler(
        entry_points=[
            CommandHandler("sotuv", sotuv_start),
            MessageHandler(filters.Text([BTN_SOTUV]), sotuv_start),
        ],
        states={
            MODEL_STEP: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_model)],
            PRODUCT: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_product)],
            QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_qty)],
            PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_price)],
        },
        fallbacks=[CommandHandler("bekor", bekor)],
    )

    xarajat_conv = ConversationHandler(
        entry_points=[
            CommandHandler("xarajat", xarajat_start),
            MessageHandler(filters.Text([BTN_XARAJAT]), xarajat_start),
        ],
        states={
            EXP_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, xarajat_name)],
            EXP_SUM: [MessageHandler(filters.TEXT & ~filters.COMMAND, xarajat_sum)],
        },
        fallbacks=[CommandHandler("bekor", bekor)],
    )

    mahsulot_conv = ConversationHandler(
        entry_points=[
            CommandHandler("mahsulot", mahsulot_start),
            MessageHandler(filters.Text([BTN_MAHSULOT]), mahsulot_start),
        ],
        states={
            NEW_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, mahsulot_name)],
            NEW_QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, mahsulot_qty)],
            NEW_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, mahsulot_price)],
        },
        fallbacks=[CommandHandler("bekor", bekor)],
    )

    app.add_handler(MessageHandler(filters.Text([BTN_QOLDIK]), qoldik))
    app.add_handler(MessageHandler(filters.Text([BTN_UNDO]), undo_handler))
    app.add_handler(CommandHandler("oxirgisini_bekor", undo_handler))
    app.add_handler(MessageHandler(filters.Text([BTN_HISOBOT]), hisobot_start))
    app.add_handler(CommandHandler("hisobot", hisobot_start))
    app.add_handler(CallbackQueryHandler(hisobot_callback, pattern=r"^rep:"))
    app.add_handler(MessageHandler(filters.Text([BTN_KAM_QOLDIQ]), kam_qoldiq_handler))
    app.add_handler(CommandHandler("kam_qoldiq", kam_qoldiq_handler))

    app.add_handler(sotuv_conv)
    app.add_handler(xarajat_conv)
    app.add_handler(mahsulot_conv)

    # Komandasiz, erkin matn bilan yozilgan sotuv/xarajatlarni tushunadigan handler.
    # Yuqoridagi conv handlerlar band bo'lmagan xabarlargagina ishlaydi.
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_free_text))

    # Har kuni soat 21:00da (Toshkent vaqti bilan) avtomatik kunlik hisobot yuboriladi.
    # Vaqtni o'zgartirish uchun Railway'da DAILY_REPORT_HOUR / DAILY_REPORT_MINUTE
    # nomli environment variable'lar qo'shish mumkin.
    if app.job_queue is not None:
        hour = int(os.environ.get("DAILY_REPORT_HOUR", "21"))
        minute = int(os.environ.get("DAILY_REPORT_MINUTE", "0"))
        report_time = datetime.time(hour=hour, minute=minute, tzinfo=ZoneInfo("Asia/Tashkent"))
        app.job_queue.run_daily(daily_report_job, time=report_time)
    else:
        print("OGOHLANTIRISH: job_queue mavjud emas — 'pip install \"python-telegram-bot[job-queue]\"' kerak.")

    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
