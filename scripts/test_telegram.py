"""Helper to discover your Telegram chat_id or send a test notification."""

import argparse
import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
os.chdir(PROJECT_ROOT)
sys.path.insert(0, str(PROJECT_ROOT))

from main import load_config
from notify.telegram_notifier import TELEGRAM_API_BASE, TelegramNotifier, _telegram_ca_bundle


def fetch_chat_ids(bot_token: str) -> int:
    import requests

    url = f"{TELEGRAM_API_BASE}/bot{bot_token.strip()}/getUpdates"
    response = requests.get(url, timeout=15, verify=_telegram_ca_bundle())
    response.raise_for_status()
    payload = response.json()
    if not payload.get("ok"):
        print(f"Telegram API returned error: {payload}")
        return 1

    seen = {}
    for update in payload.get("result", []):
        msg = (
            update.get("message")
            or update.get("edited_message")
            or update.get("channel_post")
            or {}
        )
        chat = msg.get("chat") or {}
        chat_id = chat.get("id")
        if chat_id is not None:
            label = chat.get("username") or chat.get("title") or chat.get("first_name") or "chat"
            seen[str(chat_id)] = f"{label} ({chat.get('type', 'unknown')})"

    if not seen:
        print(
            "No messages found yet. Open your bot in Telegram, press Start / send 'hi', "
            "then run this command again."
        )
        return 1

    print("Found Telegram chat(s):")
    for cid, desc in seen.items():
        print(f"  chat_id: {cid}   ({desc})")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--get-chat-id",
        metavar="BOT_TOKEN",
        help="Fetch your numeric chat_id after sending any message to your bot.",
    )
    args = parser.parse_args()

    if args.get_chat_id:
        return fetch_chat_ids(args.get_chat_id)

    config = load_config(str(PROJECT_ROOT / "config.yaml"))
    notifier = TelegramNotifier.from_config(config)
    if not notifier.enabled:
        print(
            "Telegram is not enabled in config.local.yaml. Add:\n\n"
            "telegram:\n"
            "  enabled: true\n"
            "  bot_token: \"YOUR_BOT_TOKEN\"\n"
            "  chat_id: \"YOUR_CHAT_ID\"\n"
            "  machine_name: \"Brother-PC\"\n"
            "  notify_on_success: true\n"
            "  attach_logs_on_failure: true\n"
        )
        return 1

    ok = notifier.send_message(
        f"✅ Revenue ETL Telegram test OK [{notifier.machine_name}]"
    )
    if ok:
        print(f"Test message sent successfully to chat_id={notifier.chat_id}.")
        return 0
    print("Failed to send Telegram test message. Verify bot_token, chat_id, and network access.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
