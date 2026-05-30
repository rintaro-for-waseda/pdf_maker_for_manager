# Swim Data App

水泳の測定データ入力、プッシュオフ・ダイブの表作成、PDF出力を行う静的Webアプリです。

## 使い方

GitHub Pagesで公開する場合は、`docs/` を公開フォルダにします。

```text
https://<GitHubユーザー名>.github.io/<リポジトリ名>/
```

## 更新方法

アプリ本体は `index.py` 内のHTML/JavaScriptです。変更後、次のコマンドでGitHub Pages用のHTMLを更新します。

```bash
python3 build_static.py
```

生成された `docs/index.html` をコミットしてGitHubへpushしてください。

## 保存について

入力データはブラウザの `localStorage` に保存されます。別の端末や別ブラウザへ移す場合は、画面右上のバックアップ保存/読込を使います。
