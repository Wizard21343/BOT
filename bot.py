import os
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes
from docx import Document

import os

TOKEN = os.getenv("TOKEN")

# DOCX -> TXT
def docx_to_txt(input_path, output_path):
    doc = Document(input_path)
    text = "\n".join([p.text for p in doc.paragraphs])
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

# TXT -> DOCX
def txt_to_docx(input_path, output_path):
    doc = Document()
    with open(input_path, "r", encoding="utf-8") as f:
        for line in f:
            doc.add_paragraph(line.strip())
    doc.save(output_path)

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = update.message.document
    file_name = file.file_name

    new_file = await file.get_file()
    input_path = f"input_{file_name}"
    await new_file.download_to_drive(input_path)

    if file_name.endswith(".docx"):
        output_path = input_path.replace(".docx", ".txt")
        docx_to_txt(input_path, output_path)

    elif file_name.endswith(".txt"):
        output_path = input_path.replace(".txt", ".docx")
        txt_to_docx(input_path, output_path)

    else:
        await update.message.reply_text("Отправь .docx или .txt файл")
        return

    with open(output_path, "rb") as f:
        await update.message.reply_document(f)

    os.remove(input_path)
    os.remove(output_path)

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))
    app.run_polling()

if __name__ == "__main__":
    main()
