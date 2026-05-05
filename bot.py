import os
from collections import defaultdict, deque

from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    MessageHandler,
    CommandHandler,
    ContextTypes,
    filters
)

from docx import Document


TOKEN = os.getenv("TOKEN")

ADMIN_ID = 366339367  # <-- ВСТАВЬ СЮДА СВОЙ TELEGRAM ID


# ---------------- STORAGE ----------------
user_stats = defaultdict(int)
user_history = defaultdict(lambda: deque(maxlen=5))
user_names = {}  # user_id -> name


# ---------------- CONVERTERS ----------------
def docx_to_txt(input_path, output_path):
    doc = Document(input_path)
    text = "\n".join([p.text for p in doc.paragraphs])

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)


def txt_to_docx(input_path, output_path):
    doc = Document()

    with open(input_path, "r", encoding="utf-8") as f:
        for line in f:
            doc.add_paragraph(line.strip())

    doc.save(output_path)


# ---------------- MENU ----------------
menu = ReplyKeyboardMarkup(
    [
        ["📎 Отправить файл"],
        ["📊 Статистика", "🧾 История"],
        ["ℹ️ Помощь"]
    ],
    resize_keyboard=True
)


# ---------------- START ----------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        (
            "👋 Добро пожаловать в PRO-конвертер!\n\n"
            "🛠 Возможности:\n"
            "• DOCX → TXT\n"
            "• TXT → DOCX\n"
            "• Статистика\n"
            "• История файлов\n\n"
            "📎 Просто отправь файл или нажми кнопку 👇"
        ),
        reply_markup=menu
    )


# ---------------- HELP ----------------
async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        (
            "ℹ️ ПОМОЩЬ\n\n"
            "📎 Отправь .docx или .txt файл\n"
            "⚡ Бот автоматически конвертирует\n"
        )
    )


# ---------------- USER STATS ----------------
async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_obj = update.message.from_user
    user_id = user_obj.id
    user = user_obj.first_name

    count = user_stats[user_id]

    await update.message.reply_text(
        f"📊 Статистика\n\n👤 {user}\n📎 Файлов: {count}"
    )


# ---------------- USER HISTORY ----------------
async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    hist = user_history[user_id]

    if not hist:
        await update.message.reply_text("🧾 История пуста")
        return

    text = "🧾 Последние файлы:\n\n"
    for f in hist:
        text += f"📄 {f}\n"

    await update.message.reply_text(text)


# ---------------- ADMIN PANEL ----------------
def is_admin(user_id):
    return user_id == ADMIN_ID


async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.message.from_user.id):
        await update.message.reply_text("⛔ У тебя нет доступа")
        return

    await update.message.reply_text(
        "👑 АДМИН-ПАНЕЛЬ\n\n"
        "/users - пользователи\n"
        "/allstats - вся статистика\n"
        "/logs - действия"
    )


async def users_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.message.from_user.id):
        await update.message.reply_text("⛔ Нет доступа")
        return

    if not user_names:
        await update.message.reply_text("Нет пользователей")
        return

    text = "👥 Пользователи:\n\n"
    for uid, name in user_names.items():
        text += f"{name} (ID: {uid})\n"

    await update.message.reply_text(text)


async def all_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.message.from_user.id):
        await update.message.reply_text("⛔ Нет доступа")
        return

    total = sum(user_stats.values())

    text = f"📊 Общая статистика\n\nВсего файлов: {total}\n\n"

    for uid, count in user_stats.items():
        name = user_names.get(uid, "Unknown")
        text += f"{name}: {count}\n"

    await update.message.reply_text(text)


async def logs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.message.from_user.id):
        await update.message.reply_text("⛔ Нет доступа")
        return

    text = "🧾 Последние действия:\n\n"

    for uid, files in user_history.items():
        name = user_names.get(uid, "Unknown")
        text += f"{name}:\n"
        for f in files:
            text += f"  - {f}\n"
        text += "\n"

    await update.message.reply_text(text)


# ---------------- FILE HANDLER ----------------
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = update.message.document

        if not file:
            return

        user_obj = update.message.from_user
        user_id = user_obj.id
        user = user_obj.first_name

        user_names[user_id] = user

        file_name = file.file_name

        if not (file_name.endswith(".docx") or file_name.endswith(".txt")):
            await update.message.reply_text("❌ Только .docx или .txt")
            return

        new_file = await file.get_file()
        input_path = f"input_{file_name}"
        await new_file.download_to_drive(input_path)

        await update.message.reply_text("⚡ Конвертирую...")

        if file_name.endswith(".docx"):
            output_path = input_path.replace(".docx", ".txt")
            docx_to_txt(input_path, output_path)
        else:
            output_path = input_path.replace(".txt", ".docx")
            txt_to_docx(input_path, output_path)

        with open(output_path, "rb") as f:
            await update.message.reply_document(f)

        user_stats[user_id] += 1
        user_history[user_id].append(file_name)

        os.remove(input_path)
        os.remove(output_path)

        await update.message.reply_text("✅ Готово!")

    except Exception as e:
        await update.message.reply_text(f"⚠️ Ошибка: {str(e)}")


# ---------------- BUTTON HANDLER ----------------
async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "📎 Отправить файл":
        await update.message.reply_text("📎 Отправь файл (.docx или .txt)")

    elif text == "📊 Статистика":
        await stats(update, context)

    elif text == "🧾 История":
        await history(update, context)

    elif text == "ℹ️ Помощь":
        await help_cmd(update, context)


# ---------------- MAIN ----------------
def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))

    # admin
    app.add_handler(CommandHandler("admin", admin_panel))
    app.add_handler(CommandHandler("users", users_list))
    app.add_handler(CommandHandler("allstats", all_stats))
    app.add_handler(CommandHandler("logs", logs))

    # user
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

    app.run_polling()


if __name__ == "__main__":
    main()
