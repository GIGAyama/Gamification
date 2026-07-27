# Gamification — 小学校学習記録×ゲーミフィケーション

小学校のさまざまな学習活動を児童が主体的に楽しんで行えるようにするための、
Web アプリ「まなびクエスト」のリポジトリです。

## 📁 構成

このリポジトリは、次の 2 つをセットで運用します。

| フォルダ | 内容 | 配置先 |
|---|---|---|
| **[`manabi-quest/`](./manabi-quest/)** | **まなびクエスト本体**。学習記録・ふり返り・ゲーミフィケーション・AI所見・PDF出力に加え、GIGA山学習アプリ群の**共通学習ログ（study.v1）の収集サーバー**も内蔵 | Google Apps Script（スプレッドシートにバインド） |
| **[`manabi-portal/`](./manabi-portal/)** | **学習ポータル**。まなびクエストを iframe で表示しながら、**同じ画面から学習アプリのデータを送信**できる入口ページ。児童にはこのURLを配付します | GitHub Pages（このリポジトリから公開） |

学習アプリの学習ログは `localStorage` にあり、同一オリジンのページからしか読めません。
そのため**ポータル（github.io）を親ページ、まなびクエストを iframe** にする構成にしています
（詳しくは [manabi-portal/README.md](./manabi-portal/README.md)）。

## GitHub Pages で公開する

学習ポータルはこのリポジトリからそのまま公開できます。

1. GitHub の **Settings → Pages** を開く
2. **Source** を `Deploy from a branch`、**Branch** を `main` / `/ (root)` にして Save
3. 数分後、次のURLで公開されます

   | URL | 内容 |
   |---|---|
   | `https://gigayama.github.io/Gamification/` | 入口ページ（ポータルへのリンク） |
   | `https://gigayama.github.io/Gamification/manabi-portal/` | **学習ポータル（児童に配付するURL）** |

`manabi-quest/` は Apps Script へ貼り付けるためのソースで静的サイトとしては動かないため、
[`_config.yml`](./_config.yml) で配信対象から除外しています。

### 同一オリジンについて

学習ログ（`localStorage` の `study.records.v1`）を読めるのは、学習アプリと**同一オリジン**の
ページだけです。オリジンは **スキーム＋ホスト＋ポート** で決まり、**パスは関係しません**。
`https://gigayama.github.io/Gamification/manabi-portal/` も学習アプリ本体と同じ
`https://gigayama.github.io` オリジンなので、サブパスでの公開で問題ありません。

## セットアップ

1. [`manabi-quest/README.md`](./manabi-quest/README.md) … スプレッドシートDBの作成、GASの配置、デプロイ、各種設定
2. [`manabi-portal/README.md`](./manabi-portal/README.md) … ポータルの設定と配付URLの作り方
