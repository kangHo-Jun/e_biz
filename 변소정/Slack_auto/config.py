from dotenv import load_dotenv
import os

load_dotenv()

# Slack Auth
SLACK_TOKEN = os.getenv("SLACK_TOKEN")
SLACK_COOKIE = os.getenv("SLACK_COOKIE")
SLACK_CHANNEL_ID = os.getenv("SLACK_CHANNEL_NAME")
SLACK_BOT_USER_ID = os.getenv("SLACK_USER_ID")

# Google Sheets
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL", 300))
