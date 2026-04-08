#!/usr/bin/env python3
"""
Google Drive upload script using device code flow (no browser needed on server).
"""
import urllib.request
import urllib.parse
import json
import time
import os
import sys

CLIENT_ID = os.environ.get("GDRIVE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("GDRIVE_CLIENT_SECRET", "")
SCOPE = "https://www.googleapis.com/auth/drive.file"
FOLDER_ID = "16UpQhkJw7t20wKGc1jDMh8hTPGOkG1we"
SUBFOLDER_NAME = "メルカリ_七尾様_FS資料"

def post_json(url, data):
    encoded = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=encoded)
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())

def post_api(url, headers, body_bytes, content_type):
    req = urllib.request.Request(url, data=body_bytes, method="POST")
    for k, v in headers.items():
        req.add_header(k, v)
    req.add_header("Content-Type", content_type)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())

def get_device_code():
    print("=== Google Drive 認証を開始します ===\n")
    data = {
        "client_id": CLIENT_ID,
        "scope": SCOPE,
    }
    try:
        result = post_json("https://oauth2.googleapis.com/device/code", data)
        return result
    except Exception as e:
        print(f"エラー: デバイスコード取得に失敗しました: {e}")
        print("Google Cloud ConsoleでDrive APIが有効になっているか確認してください。")
        sys.exit(1)

def poll_for_token(device_code, interval):
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "device_code": device_code,
        "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
    }
    print("認証待機中... (Ctrl+Cでキャンセル)")
    while True:
        time.sleep(interval)
        try:
            result = post_json("https://oauth2.googleapis.com/token", data)
            if "access_token" in result:
                print("✓ 認証成功!\n")
                return result
            error = result.get("error", "")
            if error == "authorization_pending":
                print(".", end="", flush=True)
            elif error == "slow_down":
                interval += 5
            else:
                print(f"\nエラー: {result}")
                sys.exit(1)
        except Exception as e:
            print(f"\n接続エラー: {e}")
            time.sleep(5)

def create_folder(access_token, folder_name, parent_id):
    url = "https://www.googleapis.com/drive/v3/files"
    headers = {"Authorization": f"Bearer {access_token}"}
    body = json.dumps({
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }).encode("utf-8")
    result = post_api(url, headers, body, "application/json")
    return result["id"]

def upload_file(access_token, file_path, folder_id):
    with open(file_path, "rb") as f:
        content = f.read()

    file_name = os.path.basename(file_path)
    metadata = json.dumps({
        "name": file_name,
        "parents": [folder_id],
    }).encode("utf-8")

    boundary = "boundary_salescore_upload"
    body = (
        f"--{boundary}\r\n"
        f"Content-Type: application/json; charset=UTF-8\r\n\r\n"
    ).encode("utf-8") + metadata + (
        f"\r\n--{boundary}\r\n"
        f"Content-Type: text/markdown\r\n\r\n"
    ).encode("utf-8") + content + (
        f"\r\n--{boundary}--"
    ).encode("utf-8")

    url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart"
    headers = {"Authorization": f"Bearer {access_token}"}
    result = post_api(url, headers, body, f"multipart/related; boundary={boundary}")
    return result

def main():
    # Step 1: Get device code
    dc = get_device_code()

    print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"以下のURLをブラウザで開いてください（MacやスマホでOK）：")
    print(f"\n  {dc['verification_url']}")
    print(f"\n次のコードを入力してください：")
    print(f"\n  {dc['user_code']}")
    print(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n")

    # Step 2: Poll for token
    token = poll_for_token(dc["device_code"], dc.get("interval", 5))
    access_token = token["access_token"]

    # Step 3: Create subfolder
    print(f"サブフォルダ「{SUBFOLDER_NAME}」を作成中...")
    sub_id = create_folder(access_token, SUBFOLDER_NAME, FOLDER_ID)
    print(f"✓ サブフォルダ作成: https://drive.google.com/drive/folders/{sub_id}\n")

    # Step 4: Upload the FS brief
    brief_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mercari_nanao_fs_brief.md")
    if not os.path.exists(brief_path):
        print(f"ファイルが見つかりません: {brief_path}")
        sys.exit(1)

    print(f"ファイルをアップロード中: {os.path.basename(brief_path)}")
    result = upload_file(access_token, brief_path, sub_id)
    print(f"✓ アップロード完了: https://drive.google.com/file/d/{result['id']}/view\n")

    print("=== 完了 ===")
    print(f"サブフォルダ: https://drive.google.com/drive/folders/{sub_id}")

if __name__ == "__main__":
    main()
