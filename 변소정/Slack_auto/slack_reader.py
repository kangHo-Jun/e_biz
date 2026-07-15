import requests
import os
from dotenv import load_dotenv

load_dotenv()

def get_messages_base(limit=10):
    token = os.getenv("SLACK_TOKEN")
    cookie = os.getenv("SLACK_COOKIE")
    channel = os.getenv("SLACK_CHANNEL_NAME")
    bot_id = os.getenv("SLACK_USER_ID")

    headers = {
        "Authorization": f"Bearer {token}",
        "Cookie": cookie
    }

    response = requests.get(
        "https://slack.com/api/conversations.history",
        headers=headers,
        params={
            "channel": channel,
            "limit": limit
        }
    )

    data = response.json()
    print(f"[디버그] API 응답: {data.get('ok')} / 오류: {data.get('error')}")

    if not data.get("ok"):
        return []

    messages = []
    for msg in data.get("messages", []):
        if msg.get("user") != bot_id:
            continue
        text = msg.get("text", "")
        if not (text.startswith(":large_blue_circle:") or
                text.startswith(":star:")):
            continue
        messages.append({"ts": msg.get("ts"), "text": text})

    return messages

def get_recent_messages(limit=10):
    return get_messages_base(limit=limit)

def get_messages():
    return get_messages_base(limit=10)
