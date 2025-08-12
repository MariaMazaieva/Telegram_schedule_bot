Telegram Schedule Bot - ERIC
This is a simply lovely Telegram bot that converts a schedule from an .xlsx Excel file into a .ics calendar file. 
The user uploads a schedule, provides a name, and the bot generates a calendar file containing all the shifts for that specific person. 
This .ics file can then be easily imported into any calendar application (Google Calendar, Apple Calendar, Outlook, etc.).

This bot is built with efficiency, processing all data in-memory without saving any user files to the server's disk.

✨ Key Features
Excel to iCal Conversion: Transforms .xlsx schedules into universal .ics calendar files.
Personalized Schedules: Generates a unique calendar for each person specified by the user.
User Interaction: Guides the user through the process step-by-step with simple commands.
Efficient In-Memory Processing: Handles all file operations in RAM, ensuring high speed and a clean server.
Deployed on Cloud: Deployed on  Render, with secure handling of API tokens.

🚀 How to Use the Bot
Any user can interact with the live bot on Telegram with these simple steps:

Find the Bot: Open Telegram and search for the bot's username: @Eric_schedule_bot.
Start the Process: Send the message convert file to the bot.
Upload Your Schedule: The bot will ask for your file. Use the paperclip icon (📎) to upload your .xlsx schedule.
Provide Your Name: Once the file is received, the bot will ask for your name. Type your full name as it appears in the schedule.
Get Your Calendar: The bot will instantly generate and send back your personal .ics calendar file. Just tap on it to import it into your phone's calendar!

🛠️ Technologies Used
Language: Python
Key Libraries:
python-telegram-bot for interacting with the Telegram Bot API.
pandas for fast and efficient Excel data processing.
ics for generating iCalendar (.ics) files.
