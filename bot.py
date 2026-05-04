import os
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


# ---------------- DOCX -> TXT ----------------
def docx_to_txt(input_path, output_path):
    doc = Document(input_path)
    text = "\n".join([p.text for p in doc.paragraphs])

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)


# ---------------- TXT -> DOCX ----------------
def txt_to_docx(input_path, output_path):
    doc = Document()

    with open(input_path, "r", encoding="utf-8") as f:
        for line in f:
            doc.add_paragraph(line.strip())

    doc.save(output_path)


# ---------------- START MENU ----------------
menu = ReplyKeyboardMarkup(
    [
        ["📎 Отправить файл"],
        ["ℹ️ Помощь"]
    ],
    resize_keyboard=True
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет!\n\n"
        "Я — конвертер файлов ⚡\n\n"
        "📌 Что я умею:\n"
        "• DOCX → TXT\n"
        "• TXT → DOCX\n\n"
        "📎 Просто отправь файл или нажми кнопку ниже 👇",
        reply_markup=menu
    )


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "ℹ️ Помощь:\n\n"
        "📎 Отправь файл (.docx или .txt)\n"
        "⚡ Я автоматически его конвертирую\n\n"
        "❗ Поддерживаются только текстовые файлы"
    )


# ---------------- FILE HANDLER ----------------
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        file = update.message.document

        if not file:
            await update.message.reply_text("📎 Просто отправь файл (.docx или .txt)")
            return

        file_name = file.file_name

        if not (file_name.endswith(".docx") or file_name.endswith(".txt")):
            await update.message.reply_text("❌ Поддерживаются только .docx и .txt")
            return

        new_file = await file.get_file()
        input_path = f"input_{file_name}"
        await new_file.download_to_drive(input_path)

        await update.message.reply_text("⚡ Конвертирую файл...")

        # DOCX -> TXT
        if file_name.endswith(".docx"):
            output_path = input_path.replace(".docx", ".txt")
            docx_to_txt(input_path, output_path)

        # TXT -> DOCX
        else:
            output_path = input_path.replace(".txt", ".docx")
            txt_to_docx(input_path, output_path)

        with open(output_path, "rb") as f:
            await update.message.reply_document(f)

        os.remove(input_path)
        os.remove(output_path)

        await update.message.reply_text("✅ Готово! Можешь отправить ещё файл 😎")

    except Exception as e:
        await update.message.reply_text(f"⚠️ Ошибка: {str(e)}")


# ---------------- TEXT BUTTONS ----------------
async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text

    if text == "📎 Отправить файл":
        await update.message.reply_text("📎 Просто отправь .docx или .txt файл")
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
