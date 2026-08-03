# GIGA Standard v4 監査：まなびクエスト（Gamification）

- 監査日: 2026-08-03
- 対象コミット: `aa9e852`
- 実施者: GIGA Standard v4 Rollout Engineer（Part III / `/rollout`）
- **Phase 0（監査）時点では、コードを1行も変更していません。**
- **その後、合意のうえで P0 → P1 → P2 → P3 → P4 を実施しました。結果は末尾の[対応後の状態](#対応後の状態)を参照してください。**

---

## 0. リポジトリの型

| 判定 | 根拠（実測） |
|---|---|
| **C+型**（GitHub Pages シェル + GAS ウェブアプリ） | `manabi-quest/*.gs` 13本 + `appsscript.json` あり。`vite.config.*` なし、`manifest_version` なし。`manabi-portal/index.html` が GitHub Pages 側のシェルとして GAS 本体を iframe 表示し、PWA・学習ログ送信・下部ナビを担当している |

```
/                        入口ページ（index.html）＋ PWA 一式（manifest.json / sw.js / offline.html / icons/）
/manabi-portal/          学習ポータル（児童に配付するURL。PWA の本体シェル）
/manabi-quest/           GAS へ貼り付けるソース（GitHub Pages では配信しない：_config.yml で exclude）
/tools/                  受信側ロジックの自動テスト（Node、package.json なし）
```

> C+型としての構造は **すでに GIGA Standard v4 の推奨形になっている**。
> Part I §3-5 の「GAS ウェブアプリは iframe 内のため PWA 化できない → GitHub Pages 側にシェルを置く」への対処が済んでいる。

---

## A. 法務・配布

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| A1 | LICENSE 実ファイル | ❌ | リポジトリ直下に存在しない | **P0** |
| A2 | `.gitignore` | ❌ | 存在しない。`.clasp.json` / `.env` を将来うっかり commit する余地がある | **P0** |
| A3 | `.github/dependabot.yml` | ❌ | `.github/` ディレクトリ自体が存在しない。npm 依存・GitHub Actions ともに現状ゼロなので実害は小さいが、`tools/` は Node 実行前提 | **P0** |
| A4 | README.md / MANUAL.md 両方 | ⚠️ | `README.md`（11KB・構成/公開手順/PWA/セットアップあり）と `manabi-quest/README.md`・`manabi-portal/README.md` は充実。**先生向けの `MANUAL.md`（専門用語ゼロ・「うまくいかないとき」節）が無い** | **P3** |

**秘密情報の混入**：`git ls-files` で `.clasp.json` / `.env` の追跡なし。ソース内にスプレッドシートID・APIキー・メールアドレスの直書きは検出されず。学習ログ送信キーは URL パラメータ経由で端末の `localStorage` に入る設計で、ソースには含まれない。**→ 履歴の書き換えが必要な事案は無い。**

---

## B. セキュリティ

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| B1 | CSP（`connect-src` が最小） | ❌ | `Content-Security-Policy` を持つ HTML は **0 件**（`index.html` / `manabi-portal/index.html` / `offline.html` / `manabi-quest/index.html` すべて） | **P1（要検証・後述の注意あり）** |
| B2 | 秘密情報・IDの直書きなし | ✅ | 検出なし（上記 A 参照） | — |
| B3 | OAuth スコープ最小 | ❌ | `appsscript.json` に **`https://www.googleapis.com/auth/drive`（Drive 全体）** あり。`https://mail.google.com/` は無し（`script.send_mail` のみ＝可） | **P0（ただし要人間判断／後述）** |
| B4 | `postMessage` の宛先が `*` でない | ⚠️ | 2 箇所で `origin \|\| '*'` のフォールバックあり<br>・`manabi-quest/js_core.html:261` → `window.top.postMessage(..., Portal.origin \|\| '*')`<br>・`manabi-portal/index.html:750` → `state.appWindow.postMessage(..., state.appOrigin \|\| '*')`<br>いずれも相手 origin が確定していれば `*` は使われないが、確定前に送るとワイルドカードになる | **P1** |
| B5 | サーバー側の多段ガード | ✅ | `Session.getActiveUser().getEmail()` を起点に `assertTeacher_()`（`03_main.gs:159`）等で役割判定。フロントの出し分けに依存していない | — |

### B3 について（**修正を保留し、人間の判断を求めます**）

Part IV は `auth/drive`（全体）を禁止しているが、このリポジトリでは実際に **全体スコープでないと動かない可能性が高い**：

| 使用箇所 | 呼び出し | `drive.file` で足りるか |
|---|---|:--:|
| `02_setup.gs:332` | `DriveApp.getFolderById(folderId)`（先生が入力した任意のフォルダID） | ❌ 足りない |
| `08_pdf.gs:49,303,352` | `DriveApp.getFileById(ss.getId()).getParents().next()`（バインド元シートの親フォルダ） | ⚠️ 不確実 |
| `09_ops.gs:264-265` | アーカイブ用スプレッドシートの `moveTo(folder)` | ⚠️ 不確実 |

`drive` → `drive.file` に落とすと、**学期末ポートフォリオPDFの保存先解決とアーカイブ移動が本番で壊れる**恐れがある。しかも変更後は全先生に**再認可**が必要で、失敗しても児童側からは原因が見えない。

→ 停止条件「たぶん大丈夫の域を出ないとき」に該当。**P0 では変更せず、次の案を提示するに留めます。**

- 案①（推奨）：`getFolderById` による任意フォルダ指定をやめ、保存先を「バインド元シートと同じフォルダ」に固定したうえで `drive.file` へ落とす。**GAS 側の機能変更を伴うため別PR・別合意**
- 案②：スコープは据え置き、README に「なぜ Drive 全体が必要か」を明記して監査時に説明できるようにする

### B1 について（CSP）

`manabi-quest/index.html` は次の外部リソースに依存している（GAS 本体側）。

| 通信先 | 用途 |
|---|---|
| `fonts.googleapis.com` / `fonts.gstatic.com` | Zen Maru Gothic |
| `cdn.jsdelivr.net` | Bootstrap 5.3.3 / Bootstrap Icons 1.11.3 / SweetAlert2 11 / kuroshiro 1.2.0 / kuroshiro-analyzer-kuromoji 1.1.0 |
| `www.gstatic.com` | Google Charts loader |

**GAS ウェブアプリは HTML の `<meta http-equiv="Content-Security-Policy">` が iframe サンドボックス（`googleusercontent.com`）側の制約と競合しやすく、ここへの CSP 投入は Part III §P1-9 の「確認できない環境なら投入せず、手順書として PR に添える」に該当する。**
→ **GAS 側（`manabi-quest/`）には CSP を入れない。** GitHub Pages 側（`index.html` / `manabi-portal/index.html` / `offline.html`）は自己完結しており外部リソースがゼロなので、こちらには**厳しい CSP を安全に投入できる**。

---

## C. 堅牢性

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| C1 | `LockService` + try/finally（GAS） | ✅ | 書き込み系で使用を確認 | — |
| C2 | 自動復旧（シート再生成） | ✅ | `02_setup.gs` にシート作成・検証あり | — |
| C3 | `pagehide` で記録確定 | ❌ | `pagehide` は **0 件**。`manabi-portal/index.html:1487` に `visibilitychange` があるが、**未送信キューの再描画のみで確定保存はしていない**。Chromebook（メモリ4GB）でタブが破棄されると、送信途中の状態が失われる | **P1** |
| C4 | 通信失敗時のリトライと明示 | ✅ | 本体経由 → 受信用URL直接POST の 2 経路フォールバックと、トーストによる明示あり | — |
| C5 | `localStorage.clear()` を使っていない | ✅ | **0 件**。`study.records.v1` にも触れていない | — |

---

## D. 表示（Part I §2）

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| D1 | viewport に `viewport-fit=cover` | ✅ | `index.html:5` / `manabi-portal/index.html:10` / `offline.html:5`、GAS 側も `03_main.gs:31` の `addMetaTag` で付与済み | — |
| D2 | `100dvh` 使用（`100vh` 単独でない） | ✅ | `index.html:29-30` / `manabi-portal/index.html:115-116` / `offline.html:12` すべてフォールバック付きの 2 行構成 | — |
| D3 | `safe-area-inset` を適用 | ✅ | 8 箇所。`--safe-t/b/l/r` を定義し、下部ナビ・シート・上部バーに反映済み | — |
| D4 | `clamp()` による fluid type | ❌ | `clamp(` は **全 HTML で 0 件**。`rem`/`px` 固定と `@media` 段組みで組んである。特に `.tab { font-size: .68rem }`（≒10.9px）は Chromebook の安価な液晶で低学年には小さい | **P1** |
| D5 | Canvas に DPR 補正 | — | `getContext(` **0 件**＝Canvas 未使用。**該当なし** | — |
| D6 | 320px 幅で横スクロールが出ない | ✅ | ポータルは `html,body { overflow: hidden }` + `max-width:100%`、入口ページは `max-width:480px` のカード。構造上、横スクロールは発生しない（実機確認は P1 の検証項目に入れる） | — |
| D7 | 画像に width/height、150KB以下 | ⚠️ | **150KB 超の PNG は 0 件**（最大 `icon-512.png` 7.6KB）＝サイズは優秀。ただし `manabi-quest/js_student.html:1446, 2464, 2644` の動的 `<img>` に **`width`/`height`/`alt` がすべて無い**（バッジ・アイテム画像。CLS の原因） | **P2** |
| D8 | コントラスト 4.5:1 以上 | ⚠️ | 本文 `#26313d` on `#eef4fb` ＝約 12:1 ✅。ただし `--muted: #7b8794` on `#eef4fb` は **約 3.3:1** で本文サイズには不足。`.tab .badge` の `#e5484d` 地に白文字は約 3.9:1 で小さい文字には不足 | **P1（別途合意のうえ実施済み）** |
| D9 | タップ領域 44px 以上・`touch-action` | ⚠️ | `.tab` は `height: 58px` ✅、`.btn` は `padding:12px 18px` で概ね 44px 超 ✅。**不足**：`.install-bar .close`（`padding:2px 6px` のみ）。`touch-action` は `.edge-zone` の 1 箇所だけで、**`touch-action: manipulation` がボタン全般に無い**（ダブルタップズームの 300ms 遅延が残る） | **P1** |
| D10 | `prefers-reduced-motion` 対応 | ⚠️ | GAS 側 `manabi-quest/css.html:1244` にはあり ✅。**GitHub Pages 側（`index.html` / `manabi-portal/index.html` / `offline.html`）に無い** — ポータルにはトースト・シートのアニメーションが 2 種類ある | **P1** |
| D11 | 提示モード（一斉授業で使う場合） | — | 児童個人が自分の記録を見るポータルであり、電子黒板での一斉提示は用途に含まれない（GAS 側に文字サイズ切替 `#font-sizer` が既にある）。**該当なし** | — |
| D12 | 印刷CSS | — | ブラウザ印刷ではなく `08_pdf.gs` の**サーバー側 PDF 生成**（学期末ポートフォリオ）で代替済み。ポータルは印刷用途を持たない。**該当なし** | — |

---

## E. PWA（Part I §3）

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| E1 | manifest の `id`/`scope`/`start_url` | ⚠️ | `"id": "/Gamification/"` は**絶対パスで正しい** ✅。`"scope": "./"` と `"start_url": "./manabi-portal/"` は相対だが、`manifest.json` がリポジトリ直下にあるため解決結果は `/Gamification/` と `/Gamification/manabi-portal/` で**実質的に正しい**。コピー元の値の残留も無し。**明示化のみ推奨（挙動は変わらない）** | **P1（表記のみ）** |
| E2 | アイコン4種 + apple-touch-icon | ✅ | 192 / 512 / maskable-192 / maskable-512 / apple-touch-icon / favicon-32 すべて実在。`icons/make_icons.py` に生成手順あり | — |
| E3 | `beforeinstallprompt` を head 最上部で捕捉 | ❌ | `manabi-portal/index.html:1324`。**HTML 全体が 71KB のうち 1324 行目**で、しかも `initPwa()` 内。校内 Wi-Fi が混雑している端末では Chrome の合図を取りこぼし、「アプリとして入れる」が出ないことがある | **P1** |
| E4 | インストールボタンをアプリ内に設置 | ✅ | 上部バーの `#install-btn` と下部の `#install-bar` の 2 系統。`isStandalone()` で起動中は非表示 ✅、iOS 向けの案内文も分岐済み ✅ | — |
| E5 | `sw.js` が自アプリ接頭辞のキャッシュのみ削除 | ✅ | `key.startsWith('manabi-')` で絞り込み済み。**他アプリを壊していない** | — |
| E6 | `sw.js` が `localStorage` に触れていない | ✅ | 0 件 | — |
| E7 | 更新通知 | ✅ | `manabi-portal/index.html:1295` で「✨ あたらしいバージョンがあります」トースト → `SKIP_WAITING` → `controllerchange` で reload。**児童向けの言葉づかいも適切** | — |
| E8 | `offline.html` | ✅ | 存在。アプリと同じ配色・フォント・`100dvh`・`viewport-fit=cover` を持つ | — |
| E9 | `APP_VERSION` を更新した | ⚠️ | `sw.js` の `const VERSION = 'v1'`。作成以来変更されていない。**リリース手順として README に明記されていない**（更新漏れは Part I §3-5 の「更新が反映されない」の主因） | **P1 + P3** |
| E10 | iOS の「ホーム画面に追加」手順を MANUAL に記載 | ⚠️ | README `## 📲 アプリとしてインストールする（PWA）` に記載あり＋アプリ内 `#install-help` でも案内 ✅。ただし **`MANUAL.md` が無い** | **P3** |

---

## F. アクセシビリティ・性能

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| F1 | alt / aria-label / aria-live | ⚠️ | `aria-label` はポータル 4 / GAS 6 箇所であり ✅。`role="dialog"` はポータル 2 箇所 ✅。**不足**：<br>・`manabi-quest/js_student.html:1446, 2464, 2644` の `<img>` に `alt` なし<br>・**ポータルに `aria-live` が 0 件** — 送信完了・失敗のトーストが読み上げられない | **P1（ポータル）/ P3（GAS）** |
| F2 | キーボードのみで全機能に到達 | ⚠️ | ボタンは `<button>` で組まれており到達可 ✅。ただし `:focus-visible` のスタイル指定がポータルに無く、**キーボード操作時に現在位置が見えない** | **P1** |
| F3 | 初回JS 300KB以下 | ❌ | GitHub Pages 側は自己完結・外部依存ゼロで **71KB（ポータル単体）** ✅。**GAS 本体は外部 CDN から Bootstrap + Bootstrap Icons + SweetAlert2 + Google Charts + kuroshiro + kuromoji を読み込む。kuroshiro の辞書だけで十数MB規模**で、300KB を大きく超える | **P2（要別途合意）** |
| F4 | 1ファイル 5,000行 / 400KB 以内 | ✅ | 最大 `manabi-quest/js_student.html` 145KB、`10_studylog.gs` 118KB。すべて 400KB 以内 | — |

### F3 について

kuroshiro（ふりがな自動付与）は児童のアクセシビリティに直結する機能で、単純に外すのは後退になる。`manabi-quest/index.html:13` のコメント通り「読み込めない環境ではルビなしで通常動作」するフォールバックは既にある。**遅延読み込み化・辞書のサブセット化は GAS 側の機能に踏み込むため、P2 として別途合意のうえ実施すべき。**

---

## G. 学習ログ `study.v1`

| # | 項目 | 判定 | 実測 | フェーズ |
|---|---|:--:|---|:--:|
| G1 | `study.v1` 準拠・個人情報を持たない | ✅ | `LOG_KEY = 'study.records.v1'`（`manabi-portal/index.html:462`）。出席番号はサーバー側のログイン判定が優先で、ポータルは氏名・メールを保持しない ✅。`localStorage.clear()` なし ✅ | — |
| G2 | 中断記録・5分ルール | ✅ | `status: 'aborted'` を受信側で正しく扱い、`10_studylog.gs:210-233` に「中断が正常な使い方であるアプリ」の一覧（`STUDY_ABORT_NORMAL_APPS`）まで用意されている。教員画面でも中断回数を表示 ✅ | — |

---

## 総括

**このリポジトリは、GIGA Standard v4 のうち構造的に難しい部分（C+型の採用・sw.js の同一オリジン安全設計・study.v1 準拠・dvh / safe-area・更新通知）を既に満たしている。** `caches.keys()` の全削除や `localStorage.clear()` といった**他アプリを壊す違反はゼロ**。

残っているのは、**入れ忘れている定型（LICENSE / .gitignore / dependabot）** と、**細部の詰め（fluid type / touch-action / reduced-motion / focus-visible / aria-live / beforeinstallprompt の位置 / pagehide）** である。

### 提案するフェーズ

| フェーズ | ブランチ | 内容 | 破壊リスク |
|---|---|---|:--:|
| **P0** | `giga-v4/p0-legal` | LICENSE（MIT）/ `.gitignore` / `.github/dependabot.yml` を新規作成。**OAuth スコープは変更しない**（B3 の理由を README に追記するに留める） | **なし** |
| **P1** | `giga-v4/p1-display-pwa` | GitHub Pages 側（`index.html` / `manabi-portal/index.html` / `offline.html`）のみを対象に：`clamp()` 化・`touch-action: manipulation`・44px 確保・`prefers-reduced-motion`・`:focus-visible`・コントラスト調整・`aria-live`・`beforeinstallprompt` を `<head>` 最上部へ・`pagehide` で確定・manifest の `scope`/`start_url` 明示化・`sw.js` の `VERSION` 更新・**GitHub Pages 側にのみ CSP 投入**（GAS 側には入れない） | **小**（GAS 本体には触れない） |
| **P2** | `giga-v4/p2-performance` | 動的 `<img>` への `width`/`height`/`alt`/`loading` 付与。**CDN 依存の自己ホスト化・kuroshiro の遅延読み込みは、GAS 本体の挙動を変えるため別途合意を要する** | 中 |
| **P3** | `giga-v4/p3-maintainability` | `MANUAL.md` 作成（先生向け・「うまくいかないとき」・iOS 手順）、README に「リリース時に `sw.js` の `VERSION` を上げる」手順と `🔐 セキュリティ設計` / `⚠️ 制限とクォータ` を追記 | なし |

### 実施しない／人間の判断を求めること

1. **OAuth スコープ `auth/drive` の縮小**（B3）— 本番の PDF・アーカイブ機能を壊す恐れ。案①②の選択をお願いします
2. **GAS 側（`manabi-quest/`）への CSP 投入**（B1）— 検証できないため手順書に留めます
3. **CDN 依存の自己ホスト化・kuroshiro の削減**（F3）— 児童のふりがな機能に影響
4. **GAS 本体の UI 文言・配色・関数名・シート列名** — Part III 絶対安全規則 5・6 により一切変更しません
5. **GitHub の `manabi-quest/` と本番 GAS の差分** — `.clasp.json` が無いため未確認。反映済みかどうかを教えてください

---

> **Phase 0 はここまでです。** 上記フェーズの実施について合意をいただいてから、P0 → P1 → … と進めます。


---

# 対応後の状態

Phase 0 の合意にもとづき **P0（法務）→ P1（表示・PWA）→ P3（保守性）** を実施しました。
**GAS 本体（`manabi-quest/`）は 1 バイトも変更していません。**
児童が見る画面の**文言・配色・アプリ名も変えていません**（Part III 絶対安全規則 5・6）。

## 解消した項目

| # | 項目 | 前 | 後 | 何をしたか |
|---|---|:--:|:--:|---|
| A1 | LICENSE | ❌ | ✅ | MIT / Copyright (c) 2026 GIGAyama |
| A2 | `.gitignore` | ❌ | ✅ | `.clasp.json` `.env` `node_modules/` `.assets-original/` ほか |
| A3 | `dependabot.yml` | ❌ | ✅ | github-actions を monthly で監視。CDN が追えない理由をコメントに明記 |
| A4 | MANUAL.md | ❌ | ✅ | 先生向け・専門用語ゼロ。「うまくいかないとき」8項目・iOS のホーム画面追加手順つき |
| B1 | CSP | ❌ | ⚠️ | `index.html` / `offline.html` に投入し実機でブロック0件を確認。**ポータルは未投入**（理由と手順は README「CSP を入れるときの手順」） |
| C3 | `pagehide` | ❌ | ✅ | 入力途中の出席番号（1〜99 のみ）を確定。学習ログ自体は送信のかたまりごとに書き戻しており元から失われない旨をコード内に明記 |
| D4 | fluid type | ❌ | ✅ | `clamp()` を 6 種導入。下部タブは 10.9px 固定 → **320px で 12.1px / 1366px で 15px**。行間 1.8 |
| D7 | `<img>` の width/height | ⚠️ | ⚠️ | 入口ページの `<img>` に付与。**GAS 側の動的 `<img>` 3 箇所は未対応（P2）** |
| D8 | コントラスト | ⚠️ | ✅ | **別途の合意のうえ実施**。下表参照 |
| D9 | タップ領域・`touch-action` | ⚠️ | ✅ | インストール案内の ✕ を 2px 余白 → **44×44px**。押せる要素すべてに `touch-action: manipulation` |
| D10 | `prefers-reduced-motion` | ⚠️ | ✅ | ポータル・入口・オフライン案内に追加（本体側は元から対応済み） |
| E1 | manifest の絶対パス | ⚠️ | ✅ | `scope` / `start_url` / shortcut を `/Gamification/…` で明示。**解決結果は従来と同一のため、インストール済みアプリに影響なし** |
| E3 | `beforeinstallprompt` の位置 | ❌ | ✅ | 1324 行目 → **`<head>` 最上部**。`pwa-installable` / `pwa-installed` イベントで受け渡し |
| E9 | `APP_VERSION` | ⚠️ | ✅ | `v1` → `v2`。上げ忘れ防止の注意書きを `sw.js` / README / MANUAL に記載 |
| E10 | iOS 手順 | ⚠️ | ✅ | MANUAL.md に手順を記載（Safari 限定である旨も明記） |
| F1 | `aria-live` | ⚠️ | ⚠️ | ポータルのトースト・未送信件数・接続テスト結果に付与。**GAS 側の `<img>` の alt 欠落は未対応（P2）** |
| F2 | フォーカス表示 | ⚠️ | ✅ | `:focus-visible { outline: 3px solid }` を 3 ファイルに追加 |

## 実機で確認したこと

`npx serve . -p 8000` + Chromium（Playwright）で自動計測しました。

| 確認項目 | 結果 |
|---|---|
| 320 / 375 / 810 / 1366px で横スクロール | **すべて発生せず**（`documentElement.scrollWidth == clientWidth`） |
| コンソールのエラー | **0 件**（`Refused to` / CSP 違反も 0 件） |
| Service Worker | `manabi-shell-v2` を生成し、シェル 10 件をキャッシュ |
| オフラインで再読み込み | **起動する**（ポータルがキャッシュから表示） |
| `pagehide` の動作 | `7` は保存され、`0`（範囲外）は保存されない |
| 下部タブの文字 | 320px で 12.1px / 1366px で 15px。58px の枠からはみ出さない |
| タップ領域 | 下部タブ 58px / ✕ 44×44px / インストール 44px |

> 入口ページの「リポジトリの README」リンクは 160×17px で 44px 未満ですが、
> **本文中のインラインリンク**であり WCAG 2.5.8 の例外に当たるため対応不要と判断しました。
>
> ポータルの `body.scrollWidth` が viewport より 60px 大きく出ますが、これは画面端の
> スワイプ目印（`.edge-hint`、`opacity:0` / `pointer-events:none`）が
> `translateX(60px)` で画面外に置かれているためです。`html { overflow: hidden }` に
> 囲まれており、実際のスクロールは発生しません（改修前からの仕様）。

## コントラストの修正（D8）

Part III 絶対安全規則 6「UI の配色を変更しない」に触れるため Phase 0 では保留していましたが、
**別途の合意を得て実施しました。** 変更したのは GitHub Pages 側の 6 か所だけで、
**GAS 本体（`manabi-quest/css.html`）の配色には触れていません。**

| 場所 | 前 | 後 | 前の比 | 後の比 |
|---|---|---|--:|--:|
| ポータル `--muted`（地 `#eef4fb`） | `#7b8794` | `#5c6874` | 3.31:1 ❌ | **5.14:1** ✅ |
| ポータル `--muted`（白地のシート・バー） | `#7b8794` | `#5c6874` | 3.66:1 ❌ | **5.70:1** ✅ |
| ポータル `.tab .badge` の地（白文字） | `#e5484d` | `#d32f2f` | 3.91:1 ❌ | **4.98:1** ✅ |
| ポータル `.tab.primary.alert` の文字（白地） | `#e5484d` | `#d32f2f` | 3.91:1 ❌ | **4.98:1** ✅ |
| 入口 `.lead` / `.note a`（白地） | `#6a7a8a` | `#5c6874` | 4.41:1 ❌ | **5.70:1** ✅ |
| 入口 `.note`（白地・注釈） | `#8a99a8` | `#626f7c` | 2.92:1 ❌ | **5.14:1** ✅ |
| オフライン案内の本文（白地） | `#6a7a8a` | `#5c6874` | 4.41:1 ❌ | **5.70:1** ✅ |

「うすい文字」「赤い警告」という見え方は保っています。**文言・レイアウト・アプリ名は変更していません。**

変更しなかったもの（測って基準を満たしていたため）：

| 場所 | 比 |
|---|--:|
| 本文 `#26313d`（地 `#eef4fb`） | 11.94:1 |
| 下部タブの文字 `#4a5763`（白地） | 7.41:1 |
| `.btn.ghost` の文字 `#46535f`（白地） | 7.89:1 |
| 入力欄のラベル `#55616e`（白地） | 6.32:1 |
| `.num-chip` `#1c5fb8`（地 `#eef4fb`） | 5.61:1 |
| `.num-chip.empty` `#c0392b`（地 `#fdecea`） | 4.76:1 |
| ダークモードの `--muted` `#93a2b3`（地 `#1a2431`） | 6.01:1 |

> `.tab:disabled { color: #bcc6d1 }`（押せない状態のタブ）は基準を下回りますが、
> WCAG 1.4.3 は**無効化されたコントロールを対象外**としているため変更していません。

## P2 — 性能とアクセシビリティ（GAS 本体を含む）

**ここから先は `manabi-quest/`（Apps Script 側）にも手を入れています。**
GitHub と Apps Script は自動同期しないため、**先生による貼り替えとデプロイの更新が必要**です
（手順は README「`manabi-quest/` を直したときは、Apps Script への貼り替えが要ります」）。
貼り替えるまでは古い本体のまま動きますが、その組み合わせでも壊れないようにしてあります。

### D7 / F1 — 動的 `<img>` の `alt` / `width` / `height`

`js_student.html` の 3 箇所を修正しました。`alt` は「読み上げとして正しいか」で決めています。

| 場所 | 内容 | `alt` | 大きさ |
|---|---|---|---|
| バッジ一覧 | 獲得したバッジの絵 | **空**（すぐ下に同じ名前が文字で出るため。読み上げの重複を避ける） | 56×56 |
| アイテム一覧 | きせかえアイテム | **アイテム名**（文字の名前が出ないため） | 68×68 |
| ガチャの結果 | 当たったアイテム | **空**（すぐ下に名前が出る） | 68×68 |

あわせて `loading="lazy"`（一覧のみ）と `decoding="async"` を付けました。
ガチャの結果は開いた瞬間に見えるので `lazy` にしていません。

### B4 — `postMessage` の宛先から `'*'` を無くした

**これが今回いちばん実害のあった項目です。**

改修前は、相手のオリジンが分かる前に送るメッセージが `'*'`（どのサイトが相手でも届ける）
になっていました。悪意のあるページがまなびクエストを iframe で埋め込むと、
画面階層（`NAV_STATE`）や起動の合図（`APP_READY`）をそのまま読めます。

| 変更 | 内容 |
|---|---|
| ポータル → 本体 | iframe のURLに `portalOrigin=<自分のオリジン>` を付けて名乗る |
| `03_main.gs` | `sanitizePortalOrigin_()` で `https://〇〇.github.io` の形だけを通し、テンプレートへ渡す |
| `index.html`（GAS） | `<body data-portal-origin="…">` として画面側に渡す |
| `js_core.html` | 起動時にこれを `Portal.origin` へ控える。`postToPortal()` は **`'*'` を一切使わない** |
| `manabi-portal/index.html` | `notifyApp()` から `\|\| '*'` を削除し、`state.appOrigin` が無ければ送らない |

**URLパラメータを改ざんされても情報は漏れません。** ブラウザは
`postMessage` の宛先オリジンが実際に一致するときだけ届けるので、
まちがった値を入れられても「届かない」だけです。

最初のあいさつ（`APP_READY`）だけは、相手がまだ名乗っていないため
`PORTAL_ORIGIN_DEFAULT`（`https://gigayama.github.io`）を使います。
これは「古いポータル ＋ 新しい本体」の組み合わせを壊さないための逃げ道でもあります。
別ドメインの GitHub Pages へ移すときは、この定数を書きかえてください
（新しいポータルなら `portalOrigin` で自動的に上書きされます）。

> 起動直後の `NAV_STATE` が送れないことがあるため、`PORTAL_READY` を受け取った時点で
> `navNotify()` を呼び直すようにしました。これが無いと、古いポータルで
> 「もどる」ボタンの状態が更新されないままになります。

### F3 — 初回に読み込む JavaScript を減らした

`<head>` に並べていた 3 つを、**必要になってから読み込む**方式に変えました。

| ライブラリ | 大きさ（実測） | 変更後 |
|---|--:|---|
| kuroshiro | 34KB | 児童画面で `initFurigana()` が呼ばれたとき。**先生の画面では読み込まない** |
| kuroshiro-analyzer-kuromoji | 77KB | 同上 |
| Google Charts ローダー | — | グラフを実際に描くとき。記録タブを開かない児童は読み込まない |

`js_core.html` に `loadScriptOnce_()` を足し、`js_student.html` の
`drawTypingChart` / `drawCalcChart` / `drawTestChart` は
「まだ用意できていなければ読み込んでから自分をもう一度呼ぶ」形にしました。
読み込めなかったときは、ふりがな無し・グラフ無しでそのまま使えます（元からその設計です）。

**残っている分**（未実施）：自前の JS が 285KB あります
（`js_core` 53KB + `js_student` 144KB + `js_teacher` 87KB）。
先生の画面と児童の画面を両方とも全員に配っている状態で、`doGet` の段階で
役割を判定して切り分ければ半分近く減らせますが、**認証まわりの構造に踏み込む**ため
今回は手を付けていません。

---

## P4 — 品質ゲートの移植

`SchoolPlan_Editor/scripts/lib/project-quality.mjs` を**正本として 1 バイトも変えずにコピー**しました
（sha256 `84bce94b…` で一致を確認）。正本が更新されたら、そのまま上書きできます。

このリポジトリは C+型で中身が 2 つに分かれているため、
`scripts/check-project.mjs`（薄い実行係）が検査を **2 回**まわします。
ライブラリ側は書きかえていません。

| 場所 | 設定 | 対象 |
|---|---|---|
| リポジトリ直下 | `quality.config.json` | GitHub Pages 側（31 ファイル） |
| `manabi-quest/` | `manabi-quest/quality.config.json` | Apps Script 側（22 ファイル） |

**結果：エラー 0 件 / 警告 0 件。**

### 明示的に許可した例外（2 件のみ）

検査を緩めるのではなく、理由を設定ファイル内に書いて許可しています。

| 例外 | 場所 | 理由 |
|---|---|---|
| `XFRAME_ALLOWALL` | `03_main.gs` | ポータルの iframe に本体を表示するために必要。無いと C+型が成り立たない |
| `auth/drive` を `forbiddenOauthScopes` から外す | `manabi-quest/quality.config.json` | 合意のうえで据え置いた判断。使用箇所ごとの理由を設定ファイルと README に記録 |

`WILDCARD_POSTMESSAGE` の例外は **ゼロ**です（B4 で解消したため）。

### 追加したテスト

`tools/check-portal-bridge.js` を新設し、`npm test` に組み込みました（24 項目）。
ポータルとの受け渡しで**宛先が `'*'` にならないこと**を、実際に関数を動かして確かめます。

- `sanitizePortalOrigin_` … `http://` / 別ドメイン / `github.io.evil.com` / パス付きを拒否
- `isPortalOrigin_` … サーバー側と判定がそろっていること
- `postToPortal` … あいさつ以外は相手不明なら送らない・宛先が `'*'` にならない
- ソース 3 ファイルに `postMessage(..., '*')` が残っていないこと

### CI

`.github/workflows/quality.yml` を追加し、push と Pull Request で
`npm run check` と `npm test` が回るようにしました。
npm パッケージを使わない構成なので `npm ci` は不要で、数秒で終わります。

---

## 実施後の全体確認

| 確認 | 結果 |
|---|---|
| `npm run check` | **0 エラー / 0 警告**（53 ファイル） |
| `npm test` | **全 6 スイート合格**（構文・ポータル受け渡し・学習ログ・課題・経験値・分析） |
| 320 / 375 / 810 / 1366px の横スクロール | 発生せず |
| コンソールのエラー | **0 件** |
| Service Worker | `manabi-shell-v4` を生成、オフライン再読み込みで起動 |
| コントラスト | 変更した全箇所が 4.5:1 以上 |

> 品質ゲートは導入直後に実際に 1 件見つけています。
> `manabi-quest/index.html` のコメント内に書いた `script` タグの文字列が
> 本物のタグとして拾われ、`HTML_SCRIPT_SYNTAX` エラーになりました（修正済み）。

---

## 残っている項目（未実施・要判断）

| # | 内容 | 状況 |
|---|---|---|
| B1 | **ポータルへの CSP 投入** | **保留（環境の制約）**。iframe が `script.google.com` → `accounts.google.com` → `〇〇.googleusercontent.com` と Google 内を多段遷移し、遷移先が端末・アカウントの状態で変わる。**本番の GAS デプロイにつないだ実機でしか検証できない**。ポリシー案と「1件でも `Refused to` が出たら入れずに戻す」手順を README に記載済み。入口ページとオフライン案内には投入済み（実機でブロック0件を確認） |
| B3 | `auth/drive` の縮小 | **据え置き（合意済み）**。理由を README と `quality.config.json` に記録 |
| F3 | 自前 JS 285KB の役割別出し分け | **未実施**。`doGet` で先生／児童を判定して `js_teacher` / `js_student` を出し分ければ半減するが、認証まわりの構造変更にあたる |
| — | `manabi-quest/` の Apps Script への反映 | **先生の操作が必要**。今回 GAS 側も変更したため、貼り替えとデプロイの更新をお願いします（手順は README） |
