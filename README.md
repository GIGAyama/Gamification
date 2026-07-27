# Gamification — 小学校学習記録×ゲーミフィケーション

小学校のさまざまな学習活動を児童が主体的に楽しんで行えるようにするための、Google Apps Script プロジェクト群です。

## 📁 構成

| フォルダ | 内容 | 状態 |
|---|---|---|
| **[`unified/`](./unified/)** | **統合版「まなびクエスト」**。3プロジェクトを1つのスプレッドシートDB・1つのWebアプリに統合。GIGA山学習アプリ群の**共通学習ログ（study.v1）の収集サーバー**も内蔵 | ✅ 推奨（最新） |
| **[`manabi-portal/`](./manabi-portal/)** | **学習ポータル**。gigayama.github.io に配置し、まなびクエストを iframe で表示しながら**同じ画面から学習アプリのデータを送信**できる入口ページ。児童にはこのURLを配付します | ✅ 推奨（最新） |
| [`study-log-sender/`](./study-log-sender/) | 学習ログ送信ページ（送信だけの単機能版）。ポータルと設定を共有します | 継続利用可 |
| `Manabi_Quest/` | 旧: ゲーミフィケーションアプリ（経験値・ガチャ・アバター・ミッション・バッジ） | 参考（旧版） |
| `assignment portfolio/` | 旧: 課題記録ポートフォリオ「学習の足あと」 | 参考（旧版） |
| `performance portfolio/` | 旧: 授業の記録・AI所見生成 | 参考（旧版） |

新規に導入する場合は **`unified/`（GAS）＋ `manabi-portal/`（GitHub Pages）** の 2 つを使用してください。
セットアップ手順・データ移行ガイドは [unified/README.md](./unified/README.md)、
ポータルの配置手順は [manabi-portal/README.md](./manabi-portal/README.md) にあります。
