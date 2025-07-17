import os
from difflib import restore
from typing import Final
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.worksheet.filters import Filters
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackContext, Updater

import pandas as pd
from datetime import datetime
from telegram import InputFile
from openpyxl import load_workbook

TOKEN: Final = '7979267305:AAEah-yqQWb2rMLAP62ksECyF9Ik0hxR51U'
BOT_USERNAME: Final = '@Ivans_schedule_bot'
EXCEL_PATH: Path = Path("Služby IK_březen  2025 2.xlsx")

# Commands
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Hello thank you for usign me! I am a very clever bot!")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("I am a schedule bot, send me your xlsx file!")


async def custom_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("This is a custom command!")

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global EXCEL_PATH

    file = update.message.document
    if not file or not file.file_name.endswith('.xlsx'):
        await update.message.reply_text("Please upload a valid .xlsx file.")
        return

    new_file = await context.bot.get_file(file.file_id)

    # Save it and update the Excel path
    download_path = Path("Experimant_bot")
    download_path.mkdir(exist_ok=True)
    EXCEL_PATH = download_path / file.file_name  # <--- Update global path
    await new_file.download_to_drive(custom_path=str(EXCEL_PATH))

    await update.message.reply_text(f"File '{file.file_name}' received and set as active schedule.")


# Responses

def handle_response(text: str) -> str:
    processed: str = text.lower()

    if 'hello' in processed:
        return 'Hey there'

    if 'how are you' in processed:
        return 'I\'m fine, thank you'

    if 'i love rohlik' in processed:
        return 'You are from Czechia'

    return 'I do not understand the command'

def create_ics_from_excel(file_path: Path, name: str, output_path: Path):
    wb = load_workbook(file_path)
    sheet = wb.active
    data = sheet.values
    columns = next(data)
    df = pd.DataFrame(data, columns=columns)

    # Clean and format
    df = df[[col for col in df.columns if isinstance(col, str) and col.lower() != "none"]]
    df = df[df["Datum"].notna()]
    df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")

    # Build calendar
    ics_lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//YourBot//Schedule Calendar//EN"
    ]

    for _, row in df.iterrows():
        date = row["Datum"]
        if pd.isna(date):
            continue
        for col in df.columns:
            if col == "Datum":
                continue
            val = row[col]
            if isinstance(val, str) and name.lower() in val.lower():
                ics_lines += [
                    "BEGIN:VEVENT",
                    f"DTSTART;VALUE=DATE:{date.strftime('%Y%m%d')}",
                    f"SUMMARY:{col} - {name}",
                    "END:VEVENT"
                ]

    ics_lines.append("END:VCALENDAR")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(ics_lines))

#

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message_type: str = update.message.chat.type
    text: str = update.message.text

    print(f'User ({update.message.from_user.id}) type: {message_type} "{text}"')

    if message_type == 'group':
        if BOT_USERNAME in text:
            new_text: str = text.replace(BOT_USERNAME, '').strip()
            response: str = handle_response(new_text)
        else:
            return
    else:
        response: str = handle_response(text)

    print('Bot: ', response)
    await update.message.reply_text(response)


async def handle_schedule_request(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    user_id = update.message.from_user.id
    print(f"User {user_id} requested schedule for: {name}")

    if not EXCEL_PATH.exists():
        await update.message.reply_text("Excel schedule file not found.")
        return

    output_dir = Path("calendar_files")
    output_dir.mkdir(exist_ok=True)  # Create the folder if it doesn't exist
    output_path = output_dir / f"{name}_schedule.ics"

    try:
        create_ics_from_excel(EXCEL_PATH, name, output_path)
        await update.message.reply_document(InputFile(str(output_path)),filename=output_path.name)
    except Exception as e:
        await update.message.reply_text(f"Failed to process schedule: {str(e)}")



async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f'Update {update} caused error {context.error}')



if __name__ == '__main__':
    print('Starting bot..')
    app = Application.builder().token(TOKEN).build()

    # Commands
    app.add_handler(CommandHandler('start', start_command))
    app.add_handler(CommandHandler('help', help_command))
    app.add_handler(CommandHandler('custom', custom_command))

    # Messages
    app.add_handler(MessageHandler(filters.TEXT, handle_message))

    # Errors
    app.add_error_handler(error)

    # File handler (for documents like PDFs, text, etc.)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_schedule_request))

    # Polls the bot
    print('Polling...')
    app.run_polling(poll_interval=3)


