#!/usr/bin/env python3
"""
Warehouse Telegram Bot v4
─────────────────────────
Серийные номера читаются через Google Cloud Vision (штрихкоды) — 100% точность.
Остальные данные (артикул, год, страна) — Claude Vision.

Два способа работы:
1. PDF — один файл содержит всё (инвойс + фото товаров)
2. Фото по отдельности — сначала инвойс, потом товары

Управление пользователями (только ADMIN_IDS):
  /adduser 123456789 Имя   – добавить
  /removeuser 123456789    – удалить
  /listusers               – список
  /myid                    – узнать свой ID
"""

import os
import io
import json
import base64
import logging
from datetime import datetime

import anthropic
import gspread
from google.oauth2.service_account import Credentials
# google vision removed
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
    WAIT_DOCUMENT,
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
        SHEET_RECEIVE: ["Дата", "Инвойс", "Оператор", "Артикул", "Ожидалось", "Получено", "Статус", "Серийный №", "Верификация", "Год произв.", "Страна происх."],
        SHEET_SHIP:    ["Дата", "Заказ/Клиент", "Оператор", "Артикул", "Ожидалось", "Отгружено", "Статус", "Серийный №", "Верификация", "Год произв.", "Страна происх."],
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


def parse_full_pdf(pdf_bytes: bytes) -> tuple[list[dict], list[dict], str]:
    """
    Читает весь PDF одним запросом через Claude.
    Возвращает (expected, found, invoice_number)
    Серийники из PDF читает Claude (штрихкоды в PDF не декодируются Vision)
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
                "- Страница 1: инвойс / delivery note с таблицей позиций\n"
                "- Остальные страницы: фото коробок с этикетками Siemens\n\n"
                "Верни ТОЛЬКО JSON объект:\n"
                "{\n"
                '  "invoice_number": "номер инвойса",\n'
                '  "invoice": [{"article": "артикул", "qty": число}, ...],\n'
                '  "goods": [{"article": "артикул", "qty": число, "serial": "серийный номер", "year": "год", "country": "страна"}, ...]\n'
                "}\n\n"
                "На этикетках Siemens:\n"
                "- артикул: поле '1P' или крупный текст типа '6ES7...'\n"
                "- серийный номер: поле 'S' — например 'S C-R3FQ8580' (бери точно как написано)\n"
                "- год: дата производства MM.YYYY\n"
                "- страна: Made in Germany = DE, Made in Vietnam = VN и т.д.\n"
                "- qty: поле QTY, если не видно — 1\n"
                "Каждая коробка = отдельная строка в goods.\n"
                "Не выдумывай данные."
            )},
        ]}],
    )
    raw = _clean_json(resp.content[0].text)
    if isinstance(raw, dict):
        return raw.get("invoice", []), raw.get("goods", []), raw.get("invoice_number", "")
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
                "Верни ТОЛЬКО JSON массив:\n"
                '[{"article": "артикул", "qty": число}, ...]\n'
                "Строки Total/Итого — пропусти."
            )},
        ]}],
    )
    return _clean_json(resp.content[0].text)


def _read_serials_pass1(img: bytes) -> list[dict]:
    """Первое чтение — читает текст поля S как обычный текст."""
    resp = _claude.messages.create(
        model="claude-opus-4-5",
        max_tokens=2000,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": _b64(img)}},
            {"type": "text", "text": (
                "Это фото коробок с товарами. Извлеки данные с каждой этикетки.\n"
                "Верни ТОЛЬКО JSON массив:\n"
                '[{"article":"...", "qty":N, "serial":"...", "year":"...", "country":"..."}, ...]\n\n'
                "ПРАВИЛА:\n"
                "- article: поле 1P, например 6ES7132-6BH01-0BA0\n"
                "- serial: текст поля S на этикетке\n"
                "  Если написано 'S C-RNA2A12Z' → serial = 'C-RNA2A12Z'\n"
                "  Если написано 'S RRT9' → serial = 'RRT9'\n"
                "  Если написано 'S LBT1137487' → serial = 'LBT1137487'\n"
                "  Читай КАЖДУЮ букву и цифру максимально внимательно\n"
                "  Если не видно чётко — пустая строка\n"
                "- year: MM.YYYY если видно, иначе пустую строку\n"
                "- country: DE/VN/CN если видно, иначе пустую строку\n"
                "- qty: поле QTY если видно, иначе 1\n"
                "Каждая коробка = отдельная строка. Не придумывай данные."
            )},
        ]}],
    )
    return _clean_json(resp.content[0].text)


def _read_serials_pass2(img: bytes, articles: list[str]) -> list[str]:
    """Второе независимое чтение — читает только серийники поля S."""
    arts_str = ", ".join(articles)
    resp = _claude.messages.create(
        model="claude-opus-4-5",
        max_tokens=1000,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": _b64(img)}},
            {"type": "text", "text": (
                f"На фото {len(articles)} коробок с артикулами: {arts_str}\n\n"
                "Для каждой коробки найди серийный номер — это текст рядом с буквой S на этикетке.\n"
                "Читай максимально внимательно каждую букву и цифру.\n"
                "Верни ТОЛЬКО JSON массив серийников в том же порядке что и коробки на фото:\n"
                '["серийник1", "серийник2", ...]\n'
                "Если серийник не виден — пустая строка.\n"
                "Не придумывай данные."
            )},
        ]}],
    )
    result = _clean_json(resp.content[0].text)
    if isinstance(result, list):
        return [str(s) for s in result]
    return [""] * len(articles)


def parse_goods_photo_claude(img: bytes) -> list[dict]:
    """
    Двойное независимое чтение серийников через Claude.
    Pass 1: читает все данные включая серийник поля S
    Pass 2: читает только серийники независимо
    Сравнивает — если совпадают = verified, если нет = предупреждение.
    """
    items = _read_serials_pass1(img)
    if not items:
        return items

    articles  = [it.get("article", "") for it in items]
    serials_2 = _read_serials_pass2(img, articles)

    for i, item in enumerate(items):
        s1 = str(item.get("serial", "")).strip().upper()
        s2 = str(serials_2[i] if i < len(serials_2) else "").strip().upper()

        if s1 and s2:
            if s1 == s2:
                item["serial_text"]    = item.get("serial", "")
                item["serial_barcode"] = serials_2[i] if i < len(serials_2) else ""
                item["verified"]       = True
                logger.info(f"Serial verified: {s1}")
            else:
                item["serial_text"]    = item.get("serial", "")
                item["serial_barcode"] = serials_2[i] if i < len(serials_2) else ""
                item["verified"]       = False
                logger.warning(f"Serial MISMATCH pass1={s1} pass2={s2}")
        else:
            item["serial_text"]    = item.get("serial", "")
            item["serial_barcode"] = serials_2[i] if i < len(serials_2) else ""
            item["verified"]       = None

    return items

def parse_goods_photo(img_bytes: bytes) -> list[dict]:
    """Двойное чтение через Claude — верификация серийников."""
    return parse_goods_photo_claude(img_bytes)


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
                "article":  art, "expected": exp, "got": got, "status": st,
                "serial":   item.get("serial",   ""),
                "verified": item.get("verified",  None),
                "year":     item.get("year",      ""),
                "country":  item.get("country",   ""),
            })

    return "\n".join(lines), has_prob, sheet_rows


# ─── Helpers ───────────────────────────────────────────────────────────────────
async def guard(update: Update) -> bool:
    if is_allowed(update.effective_user.id):
        return True
    await update.message.reply_text("⛔ Нет доступа. Обратитесь к администратору.")
    return False

async def finalize(update, context, expected, found, invoice, mode):
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
        def verified_label(v):
            if v is True:  return "✅ Оригинал"
            if v is False: return "🚨 ПОДДЕЛКА?"
            return "—"

        rows = [[now, invoice, op,
                 r["article"], r["expected"], r["got"], r["status"],
                 r["serial"], verified_label(r.get("verified")), r["year"], r["country"]]
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
        await update.message.reply_text("Использование:\n`/adduser 123456789 Имя Фамилия`", parse_mode="Markdown")
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
        f"Username: @{u.username or '—'}",
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
        "• 📄 *PDF* — бот обработает всё сам (инвойс + товары)\n"
        "• 📸 *Фото инвойса* — если нет PDF",
        parse_mode="Markdown",
    )
    return WAIT_DOCUMENT


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    doc = update.message.document
    if doc.mime_type != "application/pdf":
        await update.message.reply_text("❌ Это не PDF. Отправьте PDF файл или фото инвойса.")
        return WAIT_DOCUMENT

    msg = await update.message.reply_text("📄 Читаю PDF... подождите 15–30 секунд.")
    try:
        file      = await doc.get_file()
        pdf_bytes = bytes(await file.download_as_bytearray())
        expected, found, inv_number = parse_full_pdf(pdf_bytes)

        mode    = context.user_data.get("mode", MODE_RECEIVE)
        invoice = inv_number or "из PDF"

        if not expected:
            await msg.edit_text("❌ Не удалось найти инвойс в PDF. Попробуйте фото инвойса отдельно.")
            return WAIT_DOCUMENT

        inv_lines = [f"• `{it['article']}` — {it['qty']} шт." for it in expected]
        await msg.edit_text(
            f"📋 *Инвойс {invoice}* ({len(expected)} арт.):\n\n" + "\n".join(inv_lines) +
            f"\n\n📦 Найдено товаров: *{len(found)}* позиций\n⏳ Считаю результат...",
            parse_mode="Markdown",
        )

        await finalize(update, context, expected, found, invoice, mode)
        return ConversationHandler.END

    except Exception as e:
        logger.error(f"PDF: {e}")
        await msg.edit_text("❌ Ошибка обработки PDF. Попробуйте фото инвойса отдельно.")
        return WAIT_DOCUMENT


async def handle_invoice_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
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
        logger.warning(f"Invoice photo: {e}")
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
        "Серийники читаются через штрихкоды — 100% точность.\n"
        "Можно несколько фото. Когда закончите — /done"
    )
    return PHOTO_GOODS


async def handle_goods_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    msg = await update.message.reply_text("🔍 Читаю штрихкоды и данные...")
    try:
        photo = update.message.photo[-1]
        data  = bytes(await (await photo.get_file()).download_as_bytearray())
        items = parse_goods_photo(data)

        context.user_data["found_items"].extend(items)
        total = len(context.user_data["found_items"])

        lines      = []
        forgeries  = []
        for it in items:
            verified = it.get("verified")
            serial   = it.get("serial", "")
            s_text   = it.get("serial_text", "")
            s_bar    = it.get("serial_barcode", "")

            if verified is False:
                ico  = "🚨"
                note = f" | ⚠️ ПОДДЕЛКА? текст=`{s_text}` штрихкод=`{s_bar}`"
                forgeries.append(it.get("article","?"))
            elif verified is True:
                ico  = "✅"
                note = f" | S/N: `{serial}` ✅"
            else:
                ico  = "•"
                note = f" | S/N: `{serial}`" if serial else ""

            line = f"{ico} `{it.get('article','?')}` — {it.get('qty',1)} шт.{note}"
            if it.get("year"):    line += f" | {it['year']}"
            if it.get("country"): line += f" | {it['country']}"
            lines.append(line)

        alert = ""
        if forgeries:
            alert = f"\n\n🚨 *ВОЗМОЖНЫЕ ПОДДЕЛКИ:* {', '.join(forgeries)}"

        await msg.edit_text(
            f"✅ Это фото: {len(items)} поз.\n\n" + "\n".join(lines) +
            alert +
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

    logger.info("🤖 Warehouse bot v4 started.")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
