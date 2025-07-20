import os
from asyncio import Event
from ics import Calendar, Event

import pandas as pd
from difflib import restore
from typing import Final
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.worksheet.filters import Filters
# from pandas.conftest import names
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackContext, Updater


from datetime import datetime
from telegram import InputFile
from openpyxl import load_workbook

TOKEN: Final = '7979267305:AAEah-yqQWb2rMLAP62ksECyF9Ik0hxR51U'
BOT_USERNAME: Final = '@Ivans_schedule_bot'
EXCEL_PATH: Path = Path("Schedule.xlsx")

# Commands
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Hello thank you for usign me! I am a very clever bot!")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("I am a schedule bot, send me your xlsx file!")


async def custom_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("This is a custom command!")

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document

    if not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("This is invalid file!")
        return

    file = await context.bot.get_file(doc.file_id)
    file_name = doc.file_name

    await file.download_to_drive(custom_path="Schedule.xlsx")
    await update.message.reply_text(f"File '{file_name}' received successfully and is being processed.")

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

def create_clean_tuple(file_path: Path, name: str, output_path: Path):
    df = pd.read_excel(file_path)
    # print(df.head())
    matches = df[df.apply(lambda row: name in str(row.values), axis=1)]

    first_row = matches.iloc[0]
    first_row = first_row.drop(labels=['Den', 'Datum'], errors='ignore')  # remove if they exist
    # clean_description = "\n".join(f"{col}: {val}" for col, val in first_row.items() if pd.notna(val))
    # STEP 1: Turn the row into a list of (role, person) pairs
    assignments = [
        (role, person)
        for role, person in first_row.items()
        if pd.notna(person)
    ]

    # STEP 2: Sort however you like — e.g., move "Kuzko" to the top
    assignments.sort(key=lambda x: 0 if name in str(x[1]) else 1)

    # for item in assignments:
    #     if name in str(item[1]):
    #         assignments.remove(item)
    #         assignments.insert(1, item)  # Move to second
    #         break  # Done after first match

    # STEP 3: Print or build your string
    description = "\n".join(f"{role}: {person}" for role, person in assignments)

    print(description)
    return  assignments

def create_ics_from_excel(file_path: Path, assignments: list, output_path: Path):
    calendar = Calendar()

    user_name: str = assignments[0][1]
    print(user_name)
    # print(df.head())
    dates_for_user_name: list = get_all_dates_for_person(file_path, user_name)
    for event_date in dates_for_user_name:
        try:
            # Create the event
            event = Event()
            event.name = f"Duty Schedule - {user_name}"
            event.begin = event_date
            event.end = event_date

            # Description includes roles and names (excluding Den/Datum if needed)
            event.description = "\n".join(
                f"{role}: {person}" for role, person in assignments
                if role not in ["Den", "Datum"]
            )

            calendar.events.add(event)
        except Exception as e:
            print(f"Skipping invalid date {event_date}: {e}")
            continue

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(str(calendar))

    print(f"Calendar exported to: {output_path}")
    return

def get_all_dates_for_person(file_path: Path, user_name: str) -> list:
    df = pd.read_excel(file_path)

    # Filter rows where the person appears anywhere
    matches = df[df.apply(lambda row: user_name in str(row.values), axis=1)]

    # Extract just the 'Datum' column from those rows
    dates = matches['Datum'].dropna()

    # Convert to datetime objects (optional but clean)
    clean_dates = pd.to_datetime(dates).tolist()

    return clean_dates
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
    await update.message.reply_text("Please enter a name")
    # name: str = update.message.text
    name = update.message.text.strip()
    user_id = update.message.from_user.id
    print(f"User {user_id} requested schedule for: {name}")

    output_dir = Path("calendar_files")
    output_dir.mkdir(exist_ok=True)  # Create the folder if it doesn't exist
    output_path = output_dir / f"{name}__schedule.ics"

    if not EXCEL_PATH.exists():
        await update.message.reply_text("Excel schedule file not found.")

    clean_description: list = create_clean_tuple(EXCEL_PATH, name, output_path)
    create_ics_from_excel(EXCEL_PATH, clean_description, output_path)
    return

    # try:
    #     create_ics_from_excel(EXCEL_PATH, name, output_path)
    #     await update.message.reply_document(InputFile(str(output_path)),filename=output_path.name)
    # except Exception as e:
    #     await update.message.reply_text(f"Failed to process schedule: {str(e)}")



async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f'Update {update} caused error {context.error}')



if __name__ == '__main__':
    print('Starting bot..')
    app = Application.builder().token(TOKEN).build()

    # Commands
    # app.add_handler(CommandHandler('start', start_command))
    # app.add_handler(CommandHandler('help', help_command))
    # app.add_handler(CommandHandler('custom', custom_command))

    # Messages
    # app.add_handler(MessageHandler(filters.TEXT, handle_message))

    # Errors
    app.add_error_handler(error)

    # File handler (for documents like PDFs, text, etc.)
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_schedule_request))

    # Polls the bot
    print('Polling...')
    app.run_polling(poll_interval=3)


