#!/usr/bin/env python3
"""
Warehouse Telegram Bot v3
─────────────────────────
Два способа работы:

1. PDF (рекомендуется) — один файл содержит всё:
   /receive → выбрать режим → отправить PDF → готово

2. Фото по отдельности:
   /receive → инвойс фото → подтвердить → фото товаров → /done

Управление пользователями (только ADMIN_IDS):
  /adduser 123456789 Имя   – добавить
  /removeuser 123456789    – удалить
  /listusers               – список
  /myid                    – узнать свой ID
"""

import os
import json
import base64
import logging
from datetime import datetime

import anthropic
import gspread
from google.oauth2.service_account import Credentials
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, ConversationHandler, filters,
)

# ─── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── Config ────────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ANTHROPIC_KEY  = os.environ["ANTHROPIC_API_KEY"]
SPREADSHEET_ID = os.environ["GOOGLE_SPREADSHEET_ID"]
CREDS_FILE     = os.environ.get("GOOGLE_CREDS_FILE", "google_creds.json")

ADMIN_IDS: set[int] = {
    int(x) for x in os.environ.get("ADMIN_IDS", "").split(",") if x.strip()
}

SHEET_RECEIVE = "Приёмка"
SHEET_SHIP    = "Отгрузка"
USERS_FILE    = "allowed_users.json"

# Состояния диалога
(
    CHOOSE_MODE,
    WAIT_DOCUMENT,          # ждём PDF или фото инвойса
    CONFIRM_INVOICE_PARSE,
    PHOTO_GOODS,
) = range(4)

MODE_RECEIVE = "receive"
MODE_SHIP    = "ship"


# ─── User management ───────────────────────────────────────────────────────────
def load_users() -> dict[int, str]:
    if not os.path.exists(USERS_FILE):
        return {}
    with open(USERS_FILE, "r", encoding="utf-8") as f:
        return {int(k): v for k, v in json.load(f).items()}

def save_users(users: dict[int, str]):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump({str(k): v for k, v in users.items()}, f, ensure_ascii=False, indent=2)

def is_allowed(uid: int) -> bool:
    return uid in ADMIN_IDS or uid in load_users()

def operator_label(update: Update) -> str:
    u = update.effective_user
    if u.username:
        return f"@{u.username}"
    return u.full_name or str(u.id)


# ─── Google Sheets ──────────────────────────────────────────────────────────────
def get_worksheet(sheet_name: str):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(CREDS_FILE, scopes=scopes)
    gc = gspread.authorize(creds)
    sp = gc.open_by_key(SPREADSHEET_ID)
    headers = {
        SHEET_RECEIVE: ["Дата", "Инвойс", "Оператор", "Артикул", "Ожидалось", "Получено", "Статус", "Серийный №", "Год произв.", "Страна происх."],
        SHEET_SHIP:    ["Дата", "Заказ/Клиент", "Оператор", "Артикул", "Ожидалось", "Отгружено", "Статус", "Серийный №", "Год произв.", "Страна происх."],
    }
    try:
        ws = sp.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        ws = sp.add_worksheet(sheet_name, rows=2000, cols=10)
        ws.append_row(headers[sheet_name])
    else:
        if not ws.row_values(1):
            ws.append_row(headers[sheet_name])
    return ws

def write_rows_to_sheet(sheet_name: str, rows: list[list]):
    ws = get_worksheet(sheet_name)
    for row in rows:
        ws.append_row(row)


# ─── Claude ────────────────────────────────────────────────────────────────────
_claude = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

def _b64(data: bytes) -> str:
    return base64.standard_b64encode(data).decode()

def _clean_json(text: str):
    text = text.strip()
    if "```" in text:
        parts = text.split("```")
        text = parts[1] if len(parts) > 1 else text
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text.strip())


def parse_full_pdf(pdf_bytes: bytes) -> tuple[list[dict], list[dict]]:
    """
    Читает весь PDF одним запросом.
    Страница 1 = инвойс → expected [{article, qty}]
    Остальные страницы = этикетки/фото товаров → found [{article, qty, serial, year, country}]
    """
    resp = _claude.messages.create(
        model="claude-opus-4-5",
        max_tokens=4000,
        messages=[{"role": "user", "content": [
            {
                "type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": _b64(pdf_bytes)},
            },
            {"type": "text", "text": (
                "Этот PDF содержит:\n"
                "- Страница 1: инвойс / delivery note с таблицей позиций (артикул + количество)\n"
                "- Остальные страницы: фото коробок и товаров с этикетками\n\n"
                "Верни ТОЛЬКО JSON объект без пояснений:\n"
                "{\n"
                '  "invoice_number": "номер инвойса из страницы 1",\n'
                '  "invoice": [{"article": "артикул", "qty": число}, ...],\n'
                '  "goods": [{"article": "артикул", "qty": число, "serial": "серийный номер", "year": "год производства", "country": "страна"}, ...]\n'
                "}\n\n"
                "Правила:\n"
                "- invoice: из таблицы инвойса на странице 1, строки Total пропускай\n"
                "- goods: из этикеток на коробках (остальные страницы), каждая коробка = отдельная строка\n"
                "- serial: поле S/N, Serial Number, серийный номер с этикетки\n"
                "- year: год производства (дата выпуска на этикетке, формат YYYY или MM/YYYY)\n"
                "- country: Made in ... если видно, иначе пустая строка\n"
                "- qty всегда число, если не видно — 1\n"
                "- Только то что реально видно, не придумывай"
            )},
        ]}],
    )
    raw = _clean_json(resp.content[0].text)
    if isinstance(raw, dict):
        return (
            raw.get("invoice", []),
            raw.get("goods", []),
            raw.get("invoice_number", ""),
        )
    return [], [], ""


def parse_invoice_photo(img: bytes) -> list[dict]:
    """Читает фото инвойса → [{article, qty}]"""
    resp = _claude.messages.create(
        model="claude-opus-4-5",
        max_tokens=1500,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": _b64(img)}},
            {"type": "text", "text": (
                "Это фото инвойса или delivery note. "
                "Извлеки список позиций: артикул и количество. "
                "Верни ТОЛЬКО JSON массив без пояснений:\n"
                '[{"article": "артикул", "qty": число}, ...]\n'
                "Строки Total/Итого — пропусти. Не выдумывай данные."
            )},
        ]}],
    )
    return _clean_json(resp.content[0].text)


def parse_goods_photo(img: bytes) -> list[dict]:
    """Читает фото товаров/коробок → [{article, qty, serial, year, country}]"""
    resp = _claude.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": _b64(img)}},
            {"type": "text", "text": (
                "Это фото коробок с товарами или таблицы. Извлеки ВСЕ позиции. "
                "Верни ТОЛЬКО JSON массив:\n"
                '[{"article":"...", "qty":N, "serial":"...", "year":"...", "country":"..."}, ...]\n'
                "Total/Итого — пропусти. qty всегда число. Только то что видно."
            )},
        ]}],
    )
    return _clean_json(resp.content[0].text)


# ─── Comparison ────────────────────────────────────────────────────────────────
def compare(expected: list[dict], found: list[dict]):
    exp_map: dict[str, int] = {}
    for it in expected:
        art = str(it.get("article", "")).upper().strip()
        exp_map[art] = exp_map.get(art, 0) + int(it.get("qty", 1))

    found_map: dict[str, int] = {}
    found_details: dict[str, list] = {}
    for it in found:
        art = str(it.get("article", "")).upper().strip()
        found_map[art] = found_map.get(art, 0) + int(it.get("qty", 1))
        found_details.setdefault(art, []).append(it)

    all_arts = sorted(set(exp_map) | set(found_map))
    has_prob = False
    lines    = []
    meta     = []

    for art in all_arts:
        exp = exp_map.get(art, 0)
        got = found_map.get(art, 0)
        if got == exp:
            ico, st = "✅", "OK"
        elif got < exp:
            ico, st = "⚠️", "НЕХВАТКА"
            has_prob = True
        else:
            ico, st = "⚠️", "ИЗЛИШЕК"
            has_prob = True
        diff = got - exp
        ds = f" ({'+' if diff>0 else ''}{diff})" if diff else ""
        lines.append(f"{ico} `{art}` — ожид. {exp}, факт {got}{ds}")
        meta.append((art, exp, got, st))

    sheet_rows = []
    for art, exp, got, st in meta:
        for item in found_details.get(art, [{}]):
            sheet_rows.append({
                "article": art, "expected": exp, "got": got, "status": st,
                "serial":  item.get("serial", ""),
                "year":    item.get("year",   ""),
                "country": item.get("country",""),
            })

    return "\n".join(lines), has_prob, sheet_rows


# ─── Helpers ───────────────────────────────────────────────────────────────────
async def guard(update: Update) -> bool:
    if is_allowed(update.effective_user.id):
        return True
    await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
    return False

async def finalize(update, context, expected, found, invoice, mode):
    """Сверка + вывод отчёта + запись в Sheets."""
    op = operator_label(update)
    report, has_problem, sheet_rows = compare(expected, found)
    label  = "ПРИЁМКА" if mode == MODE_RECEIVE else "ОТГРУЗКА"
    header = f"📊 *Итог — {label}*\nДокумент: `{invoice}`\n\n"

    if has_problem:
        text = header + "🚨 *РАСХОЖДЕНИЯ ОБНАРУЖЕНЫ!*\n\n" + report
    else:
        text = header + "✅ *Всё совпадает!*\n\n" + report

    await update.message.reply_text(text, parse_mode="Markdown")

    try:
        now  = datetime.now().strftime("%Y-%m-%d %H:%M")
        sn   = SHEET_RECEIVE if mode == MODE_RECEIVE else SHEET_SHIP
        rows = [[now, invoice, op,
                 r["article"], r["expected"], r["got"], r["status"],
                 r["serial"], r["year"], r["country"]]
                for r in sheet_rows]
        write_rows_to_sheet(sn, rows)
        await update.message.reply_text(f"✅ Записано → лист «{sn}»")
    except Exception as e:
        logger.error(f"Sheets: {e}")
        await update.message.reply_text(f"⚠️ Ошибка записи в таблицу: {e}")

    context.user_data.clear()


# ─── Admin commands ─────────────────────────────────────────────────────────────
async def cmd_adduser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Только для администраторов.")
        return
    args = context.args
    if not args:
        await update.message.reply_text(
            "Использование:\n`/adduser 123456789 Имя Фамилия`\n\nID узнать: /myid",
            parse_mode="Markdown",
        )
        return
    try:
        uid = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ Первый аргумент — числовой Telegram ID.")
        return
    name  = " ".join(args[1:]) if len(args) > 1 else f"User_{uid}"
    users = load_users()
    users[uid] = name
    save_users(users)
    await update.message.reply_text(f"✅ Добавлен: *{name}* (ID: `{uid}`)", parse_mode="Markdown")

async def cmd_removeuser(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Только для администраторов.")
        return
    args = context.args
    if not args:
        await update.message.reply_text("Использование: `/removeuser 123456789`", parse_mode="Markdown")
        return
    try:
        uid = int(args[0])
    except ValueError:
        await update.message.reply_text("❌ Укажите числовой Telegram ID.")
        return
    users = load_users()
    name  = users.pop(uid, None)
    if name:
        save_users(users)
        await update.message.reply_text(f"✅ Удалён: *{name}*", parse_mode="Markdown")
    else:
        await update.message.reply_text("⚠️ Пользователь не найден.")

async def cmd_listusers(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Только для администраторов.")
        return
    users = load_users()
    if not users:
        await update.message.reply_text("Список пользователей пуст.")
        return
    lines = [f"• *{name}* — `{uid}`" for uid, name in users.items()]
    await update.message.reply_text("👥 *Сотрудники с доступом:*\n\n" + "\n".join(lines), parse_mode="Markdown")

async def cmd_myid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    u = update.effective_user
    await update.message.reply_text(
        f"🆔 Ваш Telegram ID: `{u.id}`\n"
        f"Имя: {u.full_name or '—'}\n"
        f"Username: @{u.username or '—'}\n\n"
        "Отправьте этот ID администратору для получения доступа.",
        parse_mode="Markdown",
    )


# ─── Conversation ───────────────────────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not await guard(update):
        return ConversationHandler.END
    keyboard = [[
        InlineKeyboardButton("📦 Приёмка",  callback_data=MODE_RECEIVE),
        InlineKeyboardButton("🚚 Отгрузка", callback_data=MODE_SHIP),
    ]]
    await update.message.reply_text("👋 Выберите операцию:", reply_markup=InlineKeyboardMarkup(keyboard))
    return CHOOSE_MODE


async def choose_mode(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    mode = q.data
    context.user_data.clear()
    context.user_data["mode"] = mode
    label = "Приёмка" if mode == MODE_RECEIVE else "Отгрузка"
    await q.edit_message_text(
        f"{'📦' if mode==MODE_RECEIVE else '🚚'} *{label}*\n\n"
        "Отправьте:\n"
        "• 📄 *PDF* (delivery note + фото товаров в одном файле) — бот обработает всё сам\n"
        "• 📸 *Фото инвойса* — если нет PDF",
        parse_mode="Markdown",
    )
    return WAIT_DOCUMENT


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получает PDF, читает инвойс + товары одним запросом, сразу финализирует."""
    doc = update.message.document
    if not doc.mime_type == "application/pdf":
        await update.message.reply_text("❌ Это не PDF. Отправьте PDF файл или фото инвойса.")
        return WAIT_DOCUMENT

    msg = await update.message.reply_text("📄 Читаю PDF... Это займёт 10–20 секунд.")
    try:
        file      = await doc.get_file()
        pdf_bytes = bytes(await file.download_as_bytearray())
        result    = parse_full_pdf(pdf_bytes)

        # parse_full_pdf возвращает tuple из 3 элементов
        if len(result) == 3:
            expected, found, inv_number = result
        else:
            expected, found = result
            inv_number = ""

        mode    = context.user_data.get("mode", MODE_RECEIVE)
        invoice = inv_number or context.user_data.get("invoice", "из PDF")

        if not expected:
            await msg.edit_text("❌ Не удалось найти инвойс в PDF. Попробуйте отправить фото инвойса отдельно.")
            return WAIT_DOCUMENT

        # Показываем что считали из инвойса
        inv_lines = [f"• `{it['article']}` — {it['qty']} шт." for it in expected]
        await msg.edit_text(
            f"📋 *Инвойс {invoice}* ({len(expected)} арт.):\n\n" + "\n".join(inv_lines) +
            f"\n\n📦 Найдено товаров на фото: *{len(found)}* позиций\n\n"
            "⏳ Считаю результат...",
            parse_mode="Markdown",
        )

        await finalize(update, context, expected, found, invoice, mode)
        return ConversationHandler.END

    except Exception as e:
        logger.error(f"PDF parse: {e}")
        await msg.edit_text(
            "❌ Ошибка при обработке PDF.\n\n"
            "Попробуйте:\n"
            "• Отправить PDF как файл (не как фото)\n"
            "• Или отправить фото инвойса отдельно"
        )
        return WAIT_DOCUMENT


async def handle_invoice_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Получает фото инвойса, читает позиции, просит подтвердить."""
    msg = await update.message.reply_text("🔍 Читаю инвойс...")
    try:
        photo = update.message.photo[-1]
        data  = bytes(await (await photo.get_file()).download_as_bytearray())
        items = parse_invoice_photo(data)

        if not items:
            await msg.edit_text("❌ Позиции не найдены. Попробуйте другое фото.")
            return WAIT_DOCUMENT

        context.user_data["expected"] = items
        lines = [f"• `{it['article']}` — {it['qty']} шт." for it in items]
        kb = [[
            InlineKeyboardButton("✅ Верно", callback_data="ok"),
            InlineKeyboardButton("🔄 Переснять", callback_data="retry"),
        ]]
        await msg.edit_text(
            f"📋 *Позиции из инвойса* ({len(items)} арт.):\n\n" +
            "\n".join(lines) + "\n\nВсё верно?",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(kb),
        )
        return CONFIRM_INVOICE_PARSE

    except Exception as e:
        logger.warning(f"Invoice photo parse: {e}")
        await msg.edit_text("❌ Не удалось распознать. Сделайте чёткое фото и отправьте снова.")
        return WAIT_DOCUMENT


async def confirm_invoice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    q = update.callback_query
    await q.answer()
    if q.data == "retry":
        await q.edit_message_text("📸 Отправьте новое фото инвойса.")
        return WAIT_DOCUMENT

    context.user_data["found_items"] = []
    action = "получили" if context.user_data["mode"] == MODE_RECEIVE else "отгружаете"
    await q.edit_message_text(
        f"✅ Инвойс принят.\n\n"
        f"📸 Отправьте фото товаров которые вы {action}.\n"
        "Можно несколько фото подряд.\n"
        "Когда закончите — /done"
    )
    return PHOTO_GOODS


async def handle_goods_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = await update.message.reply_text("🔍 Распознаю товары...")
    try:
        photo = update.message.photo[-1]
        data  = bytes(await (await photo.get_file()).download_as_bytearray())
        items = parse_goods_photo(data)

        context.user_data["found_items"].extend(items)
        total = len(context.user_data["found_items"])

        lines = []
        for it in items:
            line = f"• `{it.get('article','?')}` — {it.get('qty',1)} шт."
            if it.get("serial"):  line += f" | S/N: {it['serial']}"
            if it.get("year"):    line += f" | {it['year']}"
            if it.get("country"): line += f" | {it['country']}"
            lines.append(line)

        await msg.edit_text(
            f"✅ Это фото: {len(items)} поз.\n\n" + "\n".join(lines) +
            f"\n\n📊 Накоплено: *{total}* строк | Ещё фото или /done",
            parse_mode="Markdown",
        )
    except Exception as e:
        logger.warning(f"Goods photo: {e}")
        await msg.edit_text("❌ Не удалось распознать. Попробуйте снова или /done.")
    return PHOTO_GOODS


async def done(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    expected = context.user_data.get("expected", [])
    found    = context.user_data.get("found_items", [])
    invoice  = context.user_data.get("invoice", "—")
    mode     = context.user_data.get("mode", MODE_RECEIVE)

    if not found:
        await update.message.reply_text("❌ Нет данных. Отправьте хотя бы одно фото товаров.")
        return PHOTO_GOODS

    await finalize(update, context, expected, found, invoice, mode)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("❌ Отменено. /start для нового старта.")
    return ConversationHandler.END


# ─── App ───────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start",   start),
            CommandHandler("receive", start),
            CommandHandler("ship",    start),
        ],
        states={
            CHOOSE_MODE: [
                CallbackQueryHandler(choose_mode, pattern=f"^({MODE_RECEIVE}|{MODE_SHIP})$"),
            ],
            WAIT_DOCUMENT: [
                MessageHandler(filters.Document.PDF, handle_pdf),
                MessageHandler(filters.PHOTO, handle_invoice_photo),
            ],
            CONFIRM_INVOICE_PARSE: [
                CallbackQueryHandler(confirm_invoice, pattern="^(ok|retry)$"),
            ],
            PHOTO_GOODS: [
                MessageHandler(filters.PHOTO, handle_goods_photo),
                CommandHandler("done", done),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("adduser",    cmd_adduser))
    app.add_handler(CommandHandler("removeuser", cmd_removeuser))
    app.add_handler(CommandHandler("listusers",  cmd_listusers))
    app.add_handler(CommandHandler("myid",       cmd_myid))

    logger.info("🤖 Warehouse bot v3 started.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
