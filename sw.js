/* =====================================================================
 * sw.js — まなびクエスト（学習ポータル）の Service Worker
 * =====================================================================
 * 役割
 *   ① PWA としてインストールできるようにする（fetch ハンドラの存在が要件）
 *   ② ポータルのシェル（HTML/アイコン）をキャッシュし、起動を速くする
 *   ③ 通信できないときにも「オフラインです」の案内を出す
 *
 * スコープはこのファイルが置かれた場所（公開の入口＝サイトのルート）以下だけです。
 * 旧配信元の gigayama.github.io に同居していた学習アプリ（別リポジトリ＝別パス）には
 * 一切影響しません。
 *
 * 記録ハブ（records-hub.html）について
 *   学習アプリの中で iframe として開かれます。iframe の読み込みは「ページ遷移」
 *   として扱われるため、下の handleNavigate（ネットワーク優先）を通ります。
 *   直したものはすぐ届き、オフラインのときはキャッシュから開きます。
 *
 * キャッシュしないもの
 *   - GET 以外（学習ログの送信 POST など）
 *   - 別オリジン（script.google.com のまなびクエスト本体、CDN など）
 *     → まなびクエストは常に最新をサーバーから取得します
 * ===================================================================== */

'use strict';

// キャッシュを作り直したいときはこの版数を上げます。
//
// ⚠️ この行は手で直さない。tools/build-sw.mjs が SHELL_ASSETS の中身から書き換える。
//   手で上げる決まりだったころは、上げ忘れると児童の端末に前のまま表示され続けた
//   （「直したのに変わらない」の原因のほとんどがこれ）。2026-08-21 に12リポジトリで
//   同時に上げ忘れる事故が実際に起きている。
const VERSION = 'v8dd330b3'; /* __APP_VERSION__ */
const SHELL_CACHE = `manabi-shell-${VERSION}`;
const RUNTIME_CACHE = `manabi-runtime-${VERSION}`;

/** インストール時に先読みするアプリシェル */
const SHELL_ASSETS = [
  './',
  './manabi-portal/',
  // 利用規約・プライバシーの行き先を出す部品。並べておかないと、圏外で開いた
  // ときだけリンクが 1 本も出ない（行き先そのものは開けなくても、どこにあるかは
  // 見えているほうがいい）。入口とポータルの両方がこの 1 本を読む。
  './web/giga-app-links.js',
  // 記録ハブ。学習アプリの中で iframe として開かれるので、
  // オフラインのときも読めるようにしておきます（学習ログの写しが止まらないように）
  './records-hub.html',
  './offline.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-192.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(SHELL_CACHE);
    // 1つでも失敗すると addAll 全体が失敗するため、個別に入れて取りこぼしを許容する
    await Promise.all(SHELL_ASSETS.map(async url => {
      try {
        const res = await fetch(new Request(url, { cache: 'reload' }));
        if (res && res.ok) await cache.put(url, res);
      } catch (e) { /* 取得できないものは実行時にキャッシュされる */ }
    }));
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys
      .filter(key => key.startsWith('manabi-') && key !== SHELL_CACHE && key !== RUNTIME_CACHE)
      .map(key => caches.delete(key)));
    if (self.registration.navigationPreload) {
      await self.registration.navigationPreload.enable();
    }
    await self.clients.claim();
  })());
});

/** ページから「新しい版へ切り替えて」と言われたとき */
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

/**
 * ページ遷移: ネットワーク優先（内容の更新を最優先）。
 * つながらないときはキャッシュ → オフライン案内の順にフォールバックします。
 */
async function handleNavigate(event) {
  const cache = await caches.open(SHELL_CACHE);
  try {
    const preload = event.preloadResponse ? await event.preloadResponse : null;
    const response = preload || await fetch(event.request);
    // 設定URL（?app=…&key=…）は送信キーを含むのでキャッシュに残しません
    if (response && response.ok && !new URL(event.request.url).search) {
      cache.put(event.request, response.clone());
    }
    return response;
  } catch (e) {
    const cached = await cache.match(event.request, { ignoreSearch: true });
    if (cached) return cached;
    const offline = await cache.match('./offline.html');
    if (offline) return offline;
    return new Response('オフラインです。インターネットにつないでから、もう一度ひらいてください。', {
      status: 503,
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }
    });
  }
}

/** 静的ファイル: キャッシュ優先 + 裏で更新（stale-while-revalidate） */
async function handleAsset(request) {
  const cache = await caches.open(RUNTIME_CACHE);
  const cached = await cache.match(request);
  const network = fetch(request).then(response => {
    if (response && response.ok) cache.put(request, response.clone());
    return response;
  }).catch(() => null);
  return cached || (await network) || Response.error();
}

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;                     // 学習ログの POST などは素通し

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;          // 別オリジンは素通し（本体・CDN）
  if (!url.pathname.startsWith(new URL('./', self.location).pathname)) return; // スコープ外は素通し

  if (request.mode === 'navigate') {
    event.respondWith(handleNavigate(event));
    return;
  }
  event.respondWith(handleAsset(request));
});
