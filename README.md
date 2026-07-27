# Gamification — 小学校学習記録×ゲーミフィケーション

小学校のさまざまな学習活動を児童が主体的に楽しんで行えるようにするための、
Web アプリ「まなびクエスト」のリポジトリです。

## 📁 構成

このリポジトリは、次の 2 つをセットで運用します。

| フォルダ | 内容 | 配置先 |
|---|---|---|
| **[`manabi-quest/`](./manabi-quest/)** | **まなびクエスト本体**。学習記録・ふり返り・ゲーミフィケーション・AI所見・PDF出力に加え、GIGA山学習アプリ群の**共通学習ログ（study.v1）の収集サーバー**も内蔵 | Google Apps Script（スプレッドシートにバインド） |
| **[`manabi-portal/`](./manabi-portal/)** | **学習ポータル**。まなびクエストを iframe で表示しながら、**同じ画面から学習アプリのデータを送信**できる入口ページ。児童にはこのURLを配付します | GitHub Pages（`gigayama.github.io`） |

学習アプリの学習ログは `localStorage` にあり、同一オリジンのページからしか読めません。
そのため**ポータル（github.io）を親ページ、まなびクエストを iframe** にする構成にしています
（詳しくは [manabi-portal/README.md](./manabi-portal/README.md)）。

## セットアップ

1. [`manabi-quest/README.md`](./manabi-quest/README.md) … スプレッドシートDBの作成、GASの配置、デプロイ、各種設定
2. [`manabi-portal/README.md`](./manabi-portal/README.md) … ポータルの配置と配付URLの作り方
