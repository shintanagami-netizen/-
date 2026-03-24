Markdownファイルから PowerPoint (.pptx) を生成する。

## 手順

1. ユーザーが指定したMarkdownファイル（または引数 `$ARGUMENTS`）を確認する
2. 出力ファイル名はMarkdownのファイル名と同じ名前で `.pptx` 拡張子にする（例: `input.md` → `input.pptx`）
3. 以下のコマンドを実行する：

```bash
python generate_ppt.py <入力.md> <出力.pptx> logo.png
```

4. 生成完了後、出力ファイル名を報告する

## 重要

- CLAUDE.md に記載された設計仕様・カラーパレット・レイアウトルールを必ず守ること
- logo.png が存在しない場合は `logo.png` 引数を省略する
- Markdownファイルが存在しない場合はユーザーに確認する
