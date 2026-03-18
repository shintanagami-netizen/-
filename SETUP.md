# アポイントメントリマインダー セットアップ手順

**Google Apps Script（完全無料）+ Slack Bot** で動かします。

---

## 1. Slack Bot Token の取得

1. [api.slack.com/apps](https://api.slack.com/apps) を開く
2. **「Create New App」→「From scratch」**
3. アプリ名（例：`Appointment Reminder`）とワークスペースを選択して作成
4. 左メニュー **「OAuth & Permissions」→「Bot Token Scopes」** に以下を追加：
   - `chat:write`
   - `im:write`
5. ページ上部 **「Install to Workspace」** → 許可する
6. 表示される **Bot User OAuth Token（`xoxb-...`）** をコピーしておく

---

## 2. Google Apps Script の設定

1. [script.google.com](https://script.google.com) を開く
2. **「新しいプロジェクト」** をクリック
3. エディタが開いたら、`reminder.gs` の中身を全部コピーして貼り付ける
4. 1行目の `SLACK_BOT_TOKEN` を手順1でコピーしたTokenに書き換える：
   ```js
   var SLACK_BOT_TOKEN = "xoxb-xxxxxxxxxxxx"; // ← ここ
   ```
5. 上部の **「保存」**（フロッピーアイコン）をクリック

---

## 3. 動作確認

1. 関数のドロップダウンで `sendAppointmentReminder` を選択
2. **「実行」** ボタンをクリック
3. 初回のみ「Gmailへのアクセス許可」が求められるので **「許可」**
4. Slack DMに通知が届けばOK！

---

## 4. 毎朝9時に自動実行するトリガー設定

1. 左メニューの **「トリガー（時計アイコン）」** をクリック
2. 右下 **「トリガーを追加」**
3. 以下のように設定：
   - 実行する関数：`sendAppointmentReminder`
   - イベントのソース：**時間主導型**
   - 時間ベースのトリガーのタイプ：**日付ベースのタイマー**
   - 時刻：**午前9時〜10時**
4. **「保存」** をクリック

これで毎朝9時に自動でSlack DMが届きます！

---

## 動作仕様

- **実行タイミング**: 毎朝9時（Googleのサーバーで自動実行）
- **検索期間**: 当日〜4日後（金曜日実行時に月・火曜日もカバー）
- **検索キーワード（件名）**: 商談 / 打ち合わせ / アポ / MTG / ミーティング / 面談 / 訪問 / meeting / appointment
- **通知先**: Slack DM（@メンション付き）
