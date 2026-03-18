#!/usr/bin/env python3
"""
毎朝9時に当日〜2日後の商談・アポイントメントをGmailで検索し、
自分宛にリマインドメールを送るスクリプト。

セットアップ:
  pip install -r requirements.txt
  crontab -e  →  0 9 * * * /usr/bin/python3 /path/to/reminder.py

Google認証:
  Google Cloud Console で Gmail API を有効にし、
  credentials.json をこのファイルと同じディレクトリに配置してください。
"""

import os
import base64
import datetime
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ─── 設定 ────────────────────────────────────────────────
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
MY_EMAIL = "shinta.nagami@salescore.jp"
DAYS_AHEAD = 2   # 当日 + 2日後まで

SEARCH_KEYWORDS = [
    "商談", "打ち合わせ", "アポ", "MTG", "ミーティング",
    "面談", "訪問", "meeting", "appointment",
]
# ─────────────────────────────────────────────────────────


def get_gmail_service():
    """OAuth2認証してGmail APIサービスを返す。"""
    creds = None
    token_path = os.path.join(os.path.dirname(__file__), "token.json")
    creds_path = os.path.join(os.path.dirname(__file__), "credentials.json")

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def build_query(today: datetime.date, days_ahead: int) -> str:
    """Gmail検索クエリを組み立てる。"""
    end_date = today + datetime.timedelta(days=days_ahead + 1)  # beforeは翌日を指定
    keyword_part = " OR ".join(f'subject:"{kw}"' for kw in SEARCH_KEYWORDS)
    query = (
        f"({keyword_part}) "
        f"after:{today.strftime('%Y/%m/%d')} "
        f"before:{end_date.strftime('%Y/%m/%d')}"
    )
    return query


def fetch_appointments(service, query: str) -> list[dict]:
    """クエリにマッチするメールを取得し、件名・差出人・日時を返す。"""
    result = service.users().messages().list(
        userId="me", q=query, maxResults=50
    ).execute()

    messages = result.get("messages", [])
    appointments = []

    for msg in messages:
        detail = service.users().messages().get(
            userId="me", messageId=msg["id"], format="metadata",
            metadataHeaders=["Subject", "From", "Date"]
        ).execute()

        headers = {h["name"]: h["value"] for h in detail["payload"]["headers"]}
        appointments.append({
            "subject": headers.get("Subject", "(件名なし)"),
            "from":    headers.get("From", ""),
            "date":    headers.get("Date", ""),
        })

    return appointments


def format_body(today: datetime.date, appointments: list[dict]) -> str:
    """リマインドメールの本文を作成する。"""
    end_date = today + datetime.timedelta(days=DAYS_AHEAD)
    lines = [
        f"【アポイントメントリマインダー】",
        f"{today.strftime('%Y年%m月%d日')}（本日）〜 {end_date.strftime('%m月%d日')} の商談・アポ一覧",
        "=" * 50,
        "",
    ]

    if not appointments:
        lines.append("該当するアポイントメントが見つかりませんでした。")
    else:
        for i, appo in enumerate(appointments, 1):
            lines.append(f"{i}. {appo['subject']}")
            if appo["from"] and appo["from"] != MY_EMAIL:
                lines.append(f"   差出人: {appo['from']}")
            lines.append(f"   日時: {appo['date']}")
            lines.append("")

    lines += [
        "=" * 50,
        "このメールは自動送信です。",
    ]
    return "\n".join(lines)


def send_reminder(service, subject: str, body: str):
    """自分宛にリマインドメールを送信する。"""
    msg = MIMEText(body, "plain", "utf-8")
    msg["To"] = MY_EMAIL
    msg["From"] = MY_EMAIL
    msg["Subject"] = subject

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()


def main():
    today = datetime.date.today()
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}] アポイントメントリマインダー開始")

    service = get_gmail_service()
    query = build_query(today, DAYS_AHEAD)
    print(f"検索クエリ: {query}")

    appointments = fetch_appointments(service, query)
    print(f"{len(appointments)} 件のアポイントメントを検出")

    body = format_body(today, appointments)
    subject = f"【アポリマインド】{today.strftime('%m/%d')}〜{(today + datetime.timedelta(days=DAYS_AHEAD)).strftime('%m/%d')}"

    send_reminder(service, subject, body)
    print("リマインドメール送信完了")


if __name__ == "__main__":
    main()
