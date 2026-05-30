# GitHub Pages 配布手順

このプロジェクトは `docs/` をGitHub Pagesの公開フォルダとして使います。

## 更新手順

1. アプリ本体を編集する
2. 静的HTMLを生成する

```bash
python3 build_static.py
```

3. 生成された `docs/index.html` をコミットしてGitHubへpushする
4. GitHubのリポジトリ設定で Pages の公開元を `main` ブランチの `/docs` にする

## 配布URL

Pagesを有効化すると、通常は次の形式のURLになります。

```text
https://<GitHubユーザー名>.github.io/<リポジトリ名>/
```

## 注意

- Windows/Macどちらでもブラウザだけで使えます。
- 入力データは各ブラウザの `localStorage` に保存されます。
- 端末やブラウザを変えると保存データは共有されません。必要に応じて画面右上のバックアップ保存/読込を使ってください。
