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
from openpyxl import Workbook, load_workbook

from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import (
    Application,
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
PRODUCT, QTY, PRICE = range(3)
EXP_NAME, EXP_SUM = range(3, 5)
NEW_NAME, NEW_QTY, NEW_PRICE = range(5, 8)


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


def get_qoldiq_text():
    ensure_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Mahsulotlar"]
    lines = []
    for row in range(2, ws.max_row + 1):
        nomi = ws.cell(row=row, column=1).value
        if not nomi:
            continue
        qoldiq = ws.cell(row=row, column=4).value or 0
        lines.append(f"• {nomi}: {qoldiq} dona")
    if not lines:
        return "Hozircha mahsulotlar ro'yxati bo'sh. /mahsulot buyrug'i bilan qo'shing."
    return "📦 Joriy qoldiq:\n" + "\n".join(lines)


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
        await update.message.reply_text(
            f"✅ Yozildi: {mahsulot} — {miqdor:,.0f} dona x {narx:,.0f} so'm\n"
            f"Qolgan qoldiq: {qoldiq:,.0f} dona"
        )
        return

    await update.message.reply_text(
        "Tushunmadim 🙁 Quyidagicha yozib ko'ring:\n"
        "  \"mahsulot nomi sotildi 1 ta narxi 380 ming\"\n"
        "  \"xarajat benzin uchun 50 ming\"\n"
        "Yoki /sotuv, /xarajat buyrug'idan foydalaning."
    )


def xodim_ismi(update: Update) -> str:
    user = update.effective_user
    return user.full_name or (f"@{user.username}" if user.username else str(user.id))


# ============ /start ============
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    await update.message.reply_text(
        "Assalomu alaykum! Men sotuv va xarajat hisobchi botiman.\n\n"
        "/sotuv - yangi sotuvni yozish\n"
        "/xarajat - yangi xarajatni yozish\n"
        "/qoldik - qoldiqlarni ko'rish\n"
        "/mahsulot - yangi mahsulot qo'shish\n"
        "/bekor - amalni bekor qilish"
    )


# ============ /sotuv oqimi ============
async def sotuv_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return ConversationHandler.END
    await update.message.reply_text(
        "Qaysi mahsulot sotildi? (nomini yozing)", reply_markup=ReplyKeyboardRemove()
    )
    return PRODUCT


async def sotuv_product(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["mahsulot"] = update.message.text.strip()
    await update.message.reply_text("Nechta dona sotildi?")
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

    await update.message.reply_text(
        f"✅ Yozildi: {mahsulot} — {miqdor} dona x {narx:,.0f} so'm\n"
        f"Qolgan qoldiq: {qoldiq:,.0f} dona"
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
    await update.message.reply_text(f"✅ Xarajat yozildi: {context.user_data['exp_nomi']} — {summa:,.0f} so'm")
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
    await update.message.reply_text(f"✅ Mahsulot qo'shildi: {context.user_data['new_nomi']}")
    return ConversationHandler.END


# ============ /qoldik ============
async def qoldik(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("Kechirasiz, sizda bu botdan foydalanishga ruxsat yo'q.")
        return
    await update.message.reply_text(get_qoldiq_text())


# ============ /bekor ============
async def bekor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Bekor qilindi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


def main():
    ensure_excel()
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("qoldik", qoldik))

    sotuv_conv = ConversationHandler(
        entry_points=[CommandHandler("sotuv", sotuv_start)],
        states={
            PRODUCT: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_product)],
            QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_qty)],
            PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, sotuv_price)],
        },
        fallbacks=[CommandHandler("bekor", bekor)],
    )

    xarajat_conv = ConversationHandler(
        entry_points=[CommandHandler("xarajat", xarajat_start)],
        states={
            EXP_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, xarajat_name)],
            EXP_SUM: [MessageHandler(filters.TEXT & ~filters.COMMAND, xarajat_sum)],
        },
        fallbacks=[CommandHandler("bekor", bekor)],
    )

    mahsulot_conv = ConversationHandler(
        entry_points=[CommandHandler("mahsulot", mahsulot_start)],
        states={
            NEW_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, mahsulot_name)],
            NEW_QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, mahsulot_qty)],
            NEW_PRICE: [MessageHandler(filters.TEXT & ~filters.COMMAND, mahsulot_price)],
        },
        fallbacks=[CommandHandler("bekor", bekor)],
    )

    app.add_handler(sotuv_conv)
    app.add_handler(xarajat_conv)
    app.add_handler(mahsulot_conv)

    # Komandasiz, erkin matn bilan yozilgan sotuv/xarajatlarni tushunadigan handler.
    # Yuqoridagi conv handlerlar band bo'lmagan xabarlargagina ishlaydi.
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_free_text))

    print("Bot ishga tushdi...")
    app.run_polling()


if __name__ == "__main__":
    main()
