#!/usr/bin/env python3
"""Step 1: Generate OAuth URL for Google Drive authorization."""
import urllib.parse
import os

CLIENT_ID = os.environ.get("GDRIVE_CLIENT_ID", "")
REDIRECT_URI = "http://localhost"
SCOPE = "https://www.googleapis.com/auth/drive.file"

params = {
    "client_id": CLIENT_ID,
    "redirect_uri": REDIRECT_URI,
    "response_type": "code",
    "scope": SCOPE,
    "access_type": "offline",
    "prompt": "consent",
}

url = "https://accounts.google.com/o/oauth2/auth?" + urllib.parse.urlencode(params)
print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
print("以下のURLをMacのブラウザで開いてGoogleにログインしてください：")
print()
print(url)
print()
print("承認後、ブラウザが「localhost」に繋がらないエラーを表示します。")
print("その時のアドレスバーのURL（http://localhost?code=...）を")
print("そのままClaude Codeのチャットに貼り付けてください。")
print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
