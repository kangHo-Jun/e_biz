import gspread
from google.oauth2.service_account import Credentials
import os

def write_to_sheet(text):
    scopes = ["https://spreadsheets.google.com/feeds",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(
        os.path.join(os.path.dirname(__file__), "credentials.json"),
        scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(os.getenv("GOOGLE_SHEET_ID")).sheet1

    # A열 마지막 빈 행 찾기
    col_a = sheet.col_values(1)
    next_row = len(col_a) + 1

    sheet.update_cell(next_row, 1, text)
    print(f"[시트입력] 행{next_row}: {text[:30]}...")
