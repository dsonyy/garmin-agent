import json
import os
import urllib.request
import urllib.parse
import urllib.error

TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")
TELEGRAM_API_URL = "https://api.telegram.org/bot{token}"


def _call_api(method: str, params: dict, token: str | None = None) -> dict:
    token = token or TELEGRAM_BOT_TOKEN
    if not token:
        raise ValueError("TELEGRAM_BOT_TOKEN is not set")

    url = f"{TELEGRAM_API_URL.format(token=token)}/{method}"
    data = urllib.parse.urlencode(params).encode()
    req = urllib.request.Request(url, data=data)

    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        raise RuntimeError(f"Telegram API error {e.code}: {body}") from e


def send_message(
    text: str,
    chat_id: str | None = None,
    token: str | None = None,
    parse_mode: str | None = None,
    disable_notification: bool = False,
) -> dict:
    chat_id = chat_id or TELEGRAM_CHAT_ID
    if not chat_id:
        raise ValueError("TELEGRAM_CHAT_ID is not set")

    params: dict = {"chat_id": chat_id, "text": text}
    if parse_mode:
        params["parse_mode"] = parse_mode
    if disable_notification:
        params["disable_notification"] = "true"

    return _call_api("sendMessage", params, token=token)


def send_markdown(
    text: str,
    chat_id: str | None = None,
    token: str | None = None,
    disable_notification: bool = False,
) -> dict:
    return send_message(
        text,
        chat_id=chat_id,
        token=token,
        parse_mode="MarkdownV2",
        disable_notification=disable_notification,
    )


def send_html(
    text: str,
    chat_id: str | None = None,
    token: str | None = None,
    disable_notification: bool = False,
) -> dict:
    return send_message(
        text,
        chat_id=chat_id,
        token=token,
        parse_mode="HTML",
        disable_notification=disable_notification,
    )
