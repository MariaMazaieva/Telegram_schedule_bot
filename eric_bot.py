"""
A Telegram bot that converts an Excel schedule into an .ics calendar file for a specific person.
"""
from pathlib import Path
from datetime import datetime
from typing import Final, List, Tuple

import pandas as pd
import os
import io
from ics import Calendar, Event
from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

#  Constants

TOKEN = os.getenv("TELEGRAM_TOKEN")

if not TOKEN:
    raise ValueError("No TELEGRAM_TOKEN found in environment variables. Please set it before running.")

BOT_USERNAME: Final[str] = '@Eric_schedule_bot'
EXCEL_PATH: Path = Path("Schedule.xlsx")

#  State Constants for Conversation Handling
IDLE = "IDLE"
WAITING_FOR_FILE = "WAITING_FOR_FILE"
WAITING_FOR_NAME = "WAITING_FOR_NAME"

# Commands
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.setdefault("state", IDLE)
    await update.message.reply_text("Hello thank you for using me! I am a very clever bot!")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("I am a schedule bot, send me your xlsx file!")

async def custom_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("This is a custom command!")

# Responses
def handle_response(text: str) -> str | None:
    processed = text.lower().strip()
    if processed in ['hello', 'hi', 'ahoj', 'čau']:
        return 'Hello, send a "convert file" to start the process'
    if processed in ['convert the file', 'convert file']:
        return '__CONVERT__'
    return None

# File and Data Processing

def create_clean_tuple(excel_bytes: bytes, name: str) -> List[Tuple[str, str]]:
    """Finds a person's assignments from an in-memory Excel file."""
    # Used io.BytesIO to read the bytes as if it were a file on disk
    df = pd.read_excel(io.BytesIO(excel_bytes))
    name_lower = name.lower()
    def find_name_in_row(row):
        for item in row:
            if isinstance(item, str) and item.lower() == name_lower:
                return True
        return False

    matches = df[df.apply(find_name_in_row, axis=1)]
    if matches.empty:
        return []

    first_row = matches.iloc[0]
    first_row = first_row.drop(labels=['Den', 'Datum'], errors='ignore')  # remove if they exist
    assignments = [
        (role, person) for role, person in first_row.items() if pd.notna(person)
    ]
    # moves name to the top
    assignments.sort(key=lambda x: 0 if name_lower == str(x[1]).lower() else 1)

    # description = "\n".join(f"{role}: {person}" for role, person in assignments)
    # print(description)
    return  assignments

def get_all_dates_for_person(excel_bytes: bytes, user_name: str) -> List[datetime]:
    """Extracts all dates for a person from an in-memory Excel file."""
    df = pd.read_excel(io.BytesIO(excel_bytes))
    user_name_lower = user_name.lower()
    def find_user_in_row(row):
        for item in row:
            if isinstance(item, str) and item.lower() == user_name_lower:
                return True
        return False

    # Filter rows where the person appears
    matches = df[df.apply(find_user_in_row, axis=1)]
    if 'Datum' not in matches.columns or matches.empty:
        return []

    # Extract just the 'Datum' column
    dates = matches['Datum'].dropna()

    # Convert to datetime objects
    clean_dates = pd.to_datetime(dates).tolist()
    return clean_dates
    # Also valid just return pd.to_datetime(dates).tolist()

def create_ics_from_data(assignments: list, dates: list) -> str:
    """Generates an .ics calendar and returns it as a string."""
    if not assignments:
        raise ValueError("No assignment information found for this name.")
    if not dates:
        raise ValueError("No dates found for this name.")

    calendar = Calendar()
    user_name = assignments[0][1]
    description = "\n".join(f"{role}: {person}" for role, person in assignments)

    for event_date in dates:
        event = Event()
        event.name = f"Duty Schedule - {user_name}"
        event.begin = event_date.date()
        event.make_all_day()
        event.description = description
        calendar.events.add(event)

    # Return the calendar content as a string instead of writing to a file
    return str(calendar)

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

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    state = context.user_data.get("state", IDLE)

    if state != WAITING_FOR_FILE:
        await update.message.reply_text("To get starting send me please „convert file“")
        return

    if not doc or not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("It looks suspicious to me, maybe its not .xlsx type")
        return
    try:
        file = await context.bot.get_file(doc.file_id)
        # Downloads the file's content into a bytearray in memory
        excel_bytes = await file.download_as_bytearray()

        # Stores the bytes in user_data 
        context.user_data['excel_file_bytes'] = excel_bytes

        context.user_data["state"] = WAITING_FOR_NAME
        await update.message.reply_text(
            f"File '{doc.file_name}' received successfully. Now send me the name to generate the schedule."
        )
    except Exception as e:
        await update.message.reply_text(f"Sorry, I could not process the file: {e}")
        print(f"File processing error: {e}")

async def text_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles incoming text messages and manages the conversation state."""
    text = (update.message.text or "").strip()
    state = context.user_data.get("state", IDLE)

    response= handle_response(text)
    if response == '__CONVERT__':
        context.user_data["state"] = WAITING_FOR_FILE
        await update.message.reply_text("Waiting for (.xlsx).")
        return
    elif response:
        await update.message.reply_text(response)
        return

    if state == WAITING_FOR_FILE:
        await update.message.reply_text("Please send me a (.xlsx).")
        return

    if state == WAITING_FOR_NAME:
        name = text
        if not name:
            await update.message.reply_text("Send me a name please")
            return

        # Check if the excel file bytes are in memory
        excel_bytes = context.user_data.get('excel_file_bytes')
        if not excel_bytes:
            await update.message.reply_text(
                "The schedule file is missing. Please send 'convert file' and upload it again.")
            context.user_data["state"] = WAITING_FOR_FILE
            return

        try:
            assignments = create_clean_tuple(excel_bytes, name)
            if not assignments:
                await update.message.reply_text(
                    f"Could not find the name '{name}' in the schedule. Please check the spelling and try again.")
                return

            user_name = assignments[0][1]
            dates = get_all_dates_for_person(excel_bytes, user_name)

            # Generate the calendar content as a string
            calendar_string = create_ics_from_data(assignments, dates)

            # Encode the string to bytes for sending
            calendar_bytes = calendar_string.encode('utf-8')
            output_filename = f"{name.replace(' ', '_')}_schedule.ics"

            # Send the in-memory bytes as a document
            await update.message.reply_document(document=calendar_bytes, filename=output_filename)
            await update.message.reply_text("Here is your calendar! To create another, send 'convert file' again.")

        except ValueError as e:
            await update.message.reply_text(f"Could not create the calendar: {e}")
        except Exception as e:
            await update.message.reply_text(f"An unexpected error occurred. Please contact the administrator.")
            print(f"Error during calendar generation for name '{name}': {e}")
        finally:
            # IMPORTANT: Clean up user_data to free memory and reset state
            context.user_data.clear()
            context.user_data["state"] = IDLE
        return

    # Default
    await update.message.reply_text("I didnt catch it. Text „convert file“")


async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    print(f'Update {update} caused error {context.error}')

# Main Setup

def main() -> None:
    print('Starting bot...')
    app = Application.builder().token(TOKEN).build()

    # Add command handlers
    app.add_handler(CommandHandler('start', start_command))
    app.add_handler(CommandHandler('help', help_command))

    # Add message and file handlers
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_handler))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    # Add the error handler
    app.add_error_handler(error_handler)

    # Start polling
    print('Polling...')
    app.run_polling()


if __name__ == '__main__':
    main()




