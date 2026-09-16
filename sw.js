// IRON LOG Service Worker
//
// 更新方針
// - HTML（アプリ本体）は「ネットワーク優先」。電波があれば常に最新版を開き、
//   圏外やタイムアウト（4秒）のときだけ保存済みの版を表示する。
//   以前は「キャッシュ優先」だったため、Android のホーム画面アプリが古い版のまま残りやすかった。
// - アイコンなどの静的ファイルは「キャッシュ優先＋裏で更新」。
// - GitHub Pages は max-age=600 を返すため、取得時は HTTP キャッシュを使わない（cache: 'reload'）。
const APP_VERSION = 'v10';
const CACHE_NAME = 'ironlog-v10';
const NETWORK_TIMEOUT_MS = 4000;
const ASSETS = [
  './',
  './index.html',
  './iron_log_app.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(ASSETS.map((url) => new Request(url, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.map((k) => (k !== CACHE_NAME ? caches.delete(k) : null))))
      .then(() => self.clients.claim())
  );
});

// アプリから「今のバージョンは？」と聞かれたら答える
self.addEventListener('message', (e) => {
  if (e.data === 'GET_VERSION' && e.source) {
    e.source.postMessage({ type: 'SW_VERSION', version: APP_VERSION, cache: CACHE_NAME });
  }
});

function isAppPage(request) {
  if (request.mode === 'navigate') return true;
  const url = new URL(request.url);
  return url.pathname.endsWith('/') || url.pathname.endsWith('.html');
}

function networkFirst(request) {
  return caches.open(CACHE_NAME).then((cache) => {
    const network = fetch(request, { cache: 'no-cache' }).then((res) => {
      if (res && res.status === 200) cache.put(request, res.clone());
      return res;
    }).catch(() => null);
    const timeout = new Promise((resolve) => {
      setTimeout(() => resolve(null), NETWORK_TIMEOUT_MS);
    });
    return Promise.race([network, timeout]).then((res) => {
      if (res) return res;
      // 圏外・タイムアウト：保存済みの版を表示（無ければネットワークの結果を待つ）
      return cache.match(request, { ignoreSearch: true })
        .then((cached) => cached || network.then((late) => late || Response.error()));
    });
  });
}

function cacheFirst(request) {
  return caches.match(request).then((cached) => {
    const refresh = fetch(request).then((res) => {
      if (res && res.status === 200) {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
      }
      return res;
    }).catch(() => cached);
    return cached || refresh;
  });
}

self.addEventListener('fetch', (e) => {
  // Google Apps Script などの外部通信・POST はキャッシュしない
  if (e.request.method !== 'GET') return;
  if (new URL(e.request.url).origin !== self.location.origin) return;
  e.respondWith(isAppPage(e.request) ? networkFirst(e.request) : cacheFirst(e.request));
});
