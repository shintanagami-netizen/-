# アポイントメントリマインダー セットアップ手順

## 1. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

## 2. Google Cloud 認証設定

1. [Google Cloud Console](https://console.cloud.google.com/) を開く
2. プロジェクトを作成（または既存のものを使用）
3. **APIとサービス → ライブラリ** から `Gmail API` を有効化
4. **APIとサービス → 認証情報** → 「認証情報を作成」→「OAuthクライアントID」
   - アプリケーションの種類：**デスクトップアプリ**
5. ダウンロードした JSON ファイルを `credentials.json` としてこのディレクトリに保存

## 3. Slack Bot Token の設定

1. [Slack API](https://api.slack.com/apps) でアプリを作成
2. **OAuth & Permissions** → **Bot Token Scopes** に以下を追加：
   - `chat:write`
   - `im:write`
3. アプリをワークスペースにインストールして **Bot User OAuth Token**（`xoxb-...`）をコピー
4. 環境変数に設定：

```bash
export SLACK_BOT_TOKEN=xoxb-xxxxxxxxxxxx
```

cron で使う場合は crontab に直接書くか `.env` ファイルで管理してください。

## 4. 初回認証（ブラウザが開きます）

```bash
python3 reminder.py
```

初回のみブラウザでGoogleアカウントへのアクセス許可を求められます。
承認すると `token.json` が生成され、以降は自動認証されます。

## 5. cron で毎朝9時に自動実行

```bash
crontab -e
```

以下の行を追加：

```
0 9 * * * SLACK_BOT_TOKEN=xoxb-xxxxxxxxxxxx /usr/bin/python3 /path/to/reminder.py >> /path/to/reminder.log 2>&1
```

※ `/path/to/` を実際のパスに、`xoxb-xxxxxxxxxxxx` を実際のトークンに置き換えてください。

## 動作仕様

- **実行タイミング**: 毎朝9時（cron設定による）
- **検索期間**: 当日〜4日後（金曜日実行時に月・火曜日もカバー）
- **検索キーワード（件名）**: 商談 / 打ち合わせ / アポ / MTG / ミーティング / 面談 / 訪問 / meeting / appointment
- **通知先**: Slack DM（自分宛、`U09FKNYN0LD`）
