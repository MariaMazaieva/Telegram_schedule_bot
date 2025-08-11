"""
A Telegram bot that converts an Excel schedule into an .ics calendar file for a specific person.
"""
from pathlib import Path
from datetime import datetime
from typing import Final, List, Tuple

import pandas as pd
from ics import Calendar, Event
from telegram import Update, InputFile
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

#  Constants
TOKEN: Final[str] = '7979267305:AAEah-yqQWb2rMLAP62ksECyF9Ik0hxR51U'
BOT_USERNAME: Final[str] = '@Ivans_schedule_bot'
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
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    state = context.user_data.get("state", IDLE)
    if state != WAITING_FOR_FILE:
        await update.message.reply_text("To get starting send me please „convert file“")
        return

    if not doc or not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("It looks suspicious to me, maybe its not .xlsx type")
        return

    file = await context.bot.get_file(doc.file_id)
    file_name = doc.file_name
    await file.download_to_drive(custom_path="Schedule.xlsx")
    context.user_data["state"] = WAITING_FOR_NAME
    context.user_data["last_file_name"] = file_name
    await update.message.reply_text(
        f"File '{file_name}' received successfully. Now send me the name to generate the schedule."
    )

def create_clean_tuple(file_path: Path, name: str, output_path: Path):
    """
     Finds a person's schedule in the Excel file using strict, cell-by-cell matching
     and returns their assignments for that day.

     Args:
         file_path: The path to the Excel file.
         name: The name of the person to search for (case-insensitive).

     Returns:
         A list of (role, person) tuples for the found row, with the target person's
         assignment moved to the top. Returns an empty list if the name is not found.
     """
    df = pd.read_excel(file_path)
    matches = df[df.apply(lambda row: name in str(row.values), axis=1)]
    if matches.empty:
        return []
    first_row = matches.iloc[0]
    first_row = first_row.drop(labels=['Den', 'Datum'], errors='ignore')  # remove if they exist
    assignments = [
        (role, person) for role, person in first_row.items() if pd.notna(person)
    ]
    # moves name to the top
    assignments.sort(key=lambda x: 0 if name in str(x[1]) else 1)

    description = "\n".join(f"{role}: {person}" for role, person in assignments)
    print(description)
    return  assignments

def get_all_dates_for_person(file_path: Path, user_name: str) -> list:
    """
    Extracts all dates for a specific person from the Excel file using strict matching.

    Args:
        file_path: The path to the Excel file.
        user_name: The name of the person to find dates for.

    Returns:
        A list of datetime objects for the given person.
    """
    df = pd.read_excel(file_path)

    # Filter rows where the person appears
    matches = df[df.apply(lambda row: user_name in str(row.values), axis=1)]

    # Extract just the 'Datum' column
    dates = matches['Datum'].dropna()

    # Convert to datetime objects
    clean_dates = pd.to_datetime(dates).tolist()
    return clean_dates
    # Also valid just return pd.to_datetime(dates).tolist()

def create_ics_from_excel(file_path: Path, assignments: list, output_path: Path):
    """
    Generates an .ics calendar file from a list of assignments and dates.

    Args:
        assignments: A list of (role, person) tuples.
        dates: A list of dates for the events.
        output_path: The path to save the generated .ics file.

    Raises:
        ValueError: If assignments or dates are empty.
    """
    calendar = Calendar()
    if not assignments:
        raise ValueError("There is no information for this name")
    user_name: str = assignments[0][1]
    print(user_name)
    dates_for_user_name: list = get_all_dates_for_person(file_path, user_name)
    if not dates_for_user_name:
        raise ValueError("There is no dates for this name")

    for event_date in dates_for_user_name:
        # Creates the event
        event = Event()
        event.name = f"Duty Schedule - {user_name}"
        event.begin = event_date
        event.end = event_date

        # Description includes roles and names
        event.description = "\n".join(
            f"{role}: {person}" for role, person in assignments
            if role not in ["Den", "Datum"]
        )

        calendar.events.add(event)

    # output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(str(calendar))

    print(f"Calendar exported to: {output_path}")
    return



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

        if not EXCEL_PATH.exists():
            await update.message.reply_text("File (Schedule.xlsx) is missing. Send it again please.")
            context.user_data["state"] = WAITING_FOR_FILE
            return
        try:
            out_dir = Path("calendar_files")
            out_path = out_dir / f"{name}_schedule.ics"
            assignments = create_clean_tuple(EXCEL_PATH, name, out_path)
            create_ics_from_excel(EXCEL_PATH, assignments, out_path)

            await update.message.reply_document(InputFile(str(out_path)), filename=out_path.name)
            await update.message.reply_text("Ready :)")
            context.user_data["state"] = IDLE
            context.user_data.pop("last_file_name", None)

        except Exception as e:
            await update.message.reply_text(f"Couldnt create a calendar: {e}")
        return

    # Default
    await update.message.reply_text("I didnt catch it. Text „convert file“")

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
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


