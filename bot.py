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


# ---------------- STORAGE ----------------
user_stats = defaultdict(int)
user_history = defaultdict(lambda: deque(maxlen=5))


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
            "📎 Просто отправь файл или нажми кнопку 👇\n\n"
            "👨‍💻 Создатель: @YOUR_USERNAME"
        ),
        reply_markup=menu
    )


# ---------------- HELP ----------------
async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        (
            "ℹ️ ПОМОЩЬ\n\n"
            "📎 Как пользоваться:\n"
            "• Отправь .docx или .txt файл\n"
            "• Или нажми кнопку 📎\n\n"
            "⚡ Бот автоматически конвертирует файл\n\n"
            "👨‍💻 Создатель: @YOUR_USERNAME"
        )
    )


# ---------------- STATS ----------------
async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.first_name
    count = user_stats[user]

    await update.message.reply_text(
        f"📊 Статистика\n\n"
        f"👤 {user}\n"
        f"📎 Файлов обработано: {count}"
    )


# ---------------- HISTORY ----------------
async def history(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.message.from_user.first_name
    hist = user_history[user]

    if not hist:
        await update.message.reply_text("🧾 История пуста")
        return

    text = "🧾 Последние файлы:\n\n"
    for f in hist:
        text += f"📄 {f}\n"

    await update.message.reply_text(text)


# ---------------- FILE HANDLER ----------------
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = update.message.document
        user = update.message.from_user.first_name

        if not file:
            return

        file_name = file.file_name

        if not (file_name.endswith(".docx") or file_name.endswith(".txt")):
            await update.message.reply_text("❌ Только .docx или .txt")
            return

        new_file = await file.get_file()
        input_path = f"input_{file_name}"
        await new_file.download_to_drive(input_path)

        await update.message.reply_text("⚡ Конвертирую файл...")

        # DOCX → TXT
        if file_name.endswith(".docx"):
            output_path = input_path.replace(".docx", ".txt")
            docx_to_txt(input_path, output_path)

        # TXT → DOCX
        else:
            output_path = input_path.replace(".txt", ".docx")
            txt_to_docx(input_path, output_path)

        with open(output_path, "rb") as f:
            await update.message.reply_document(f)

        # stats
        user_stats[user] += 1
        user_history[user].append(file_name)

        os.remove(input_path)
        os.remove(output_path)

        await update.message.reply_text("✅ Готово!")

    except Exception as e:
        await update.message.reply_text(f"⚠️ Ошибка: {str(e)}")


# ---------------- BUTTONS ----------------
async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "📎 Отправить файл":
        await update.message.reply_text("📎 Просто отправь файл (.docx или .txt)")

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

    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))

    app.run_polling()


if __name__ == "__main__":
    main()
