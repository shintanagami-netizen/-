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

## 3. 初回認証（ブラウザが開きます）

```bash
python3 reminder.py
```

初回のみブラウザでGoogleアカウントへのアクセス許可を求められます。
承認すると `token.json` が生成され、以降は自動認証されます。

## 4. cron で毎朝9時に自動実行

```bash
crontab -e
```

以下の行を追加：

```
0 9 * * * /usr/bin/python3 /path/to/reminder.py >> /path/to/reminder.log 2>&1
```

※ `/path/to/` を実際のパスに置き換えてください。

## 動作仕様

- **実行タイミング**: 毎朝9時（cron設定による）
- **検索期間**: 当日〜4日後（金曜日実行時に月・火曜日もカバー）
- **検索キーワード（件名）**: 商談 / 打ち合わせ / アポ / MTG / ミーティング / 面談 / 訪問 / meeting / appointment
- **送信先**: `shinta.nagami@salescore.jp`（自分宛）
