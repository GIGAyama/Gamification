# Gamification — 小学校学習記録×ゲーミフィケーション

小学校のさまざまな学習活動を児童が主体的に楽しんで行えるようにするための、Google Apps Script プロジェクト群です。

## 📁 構成

| フォルダ | 内容 | 状態 |
|---|---|---|
| **[`unified/`](./unified/)** | **統合版「まなびクエスト」**。3プロジェクトを1つのスプレッドシートDB・1つのWebアプリに統合。GIGA山学習アプリ群の**共通学習ログ（study.v1）の収集サーバー**も内蔵 | ✅ 推奨（最新） |
| **[`study-log-sender/`](./study-log-sender/)** | **学習ログ送信ページ**。gigayama.github.io に配置し、学習アプリが端末に保存した study.v1 ログをまなびクエストへ送信 | ✅ 推奨（最新） |
| `Manabi_Quest/` | 旧: ゲーミフィケーションアプリ（経験値・ガチャ・アバター・ミッション・バッジ） | 参考（旧版） |
| `assignment portfolio/` | 旧: 課題記録ポートフォリオ「学習の足あと」 | 参考（旧版） |
| `performance portfolio/` | 旧: 授業の記録・AI所見生成 | 参考（旧版） |

新規に導入する場合は **`unified/` のみ**を使用してください。セットアップ手順・データ移行ガイドは [unified/README.md](./unified/README.md) にあります。
