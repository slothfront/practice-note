// PRACTICE NOTE：ネットが無くても開けるようにする仕組み（サービスワーカー）
// - ページ本体：ネットがあれば最新を取りに行き、無い・遅いときは前回の控えを出す
// - アイコン・manifest：控えを先に出す
// - 字（Google Fonts）：一度読めたものは控えておく
// 記録（localStorage）と、GASへの送受信には触らない。
const CACHE = 'pnote-v1';
const PAGE = './';
const PRECACHE = ['./', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'apple-touch-icon.png', 'favicon-32.png'];
const WAIT_MS = 3000;

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(PRECACHE.map((u) => new Request(u, { cache: 'reload' }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

function pageFromNetwork(req) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('timeout')), WAIT_MS);
    fetch(req, { cache: 'no-cache' }).then((res) => {
      clearTimeout(timer);
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then((c) => c.put(PAGE, copy)); }
      resolve(res);
    }, (err) => { clearTimeout(timer); reject(err); });
  });
}

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  if (req.mode === 'navigate' && url.origin === location.origin) {
    e.respondWith(
      pageFromNetwork(req).catch(() =>
        caches.match(PAGE).then((hit) => hit || caches.match(req, { ignoreSearch: true }))
          .then((hit) => hit || new Response(
            '<meta charset="utf-8"><p style="font:16px sans-serif;padding:24px">一度ネットにつないで開いてください。</p>',
            { headers: { 'Content-Type': 'text/html; charset=utf-8' } }))
      )
    );
    return;
  }

  // 新しい版の確かめ（index.html の ?v=…）は控えを使わず、いつもネットに聞く
  if (url.origin === location.origin && url.searchParams.has('v')) return;

  if (url.origin === location.origin) {
    e.respondWith(caches.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req)));
    return;
  }

  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    e.respondWith(
      caches.open(CACHE).then((c) =>
        c.match(req).then((hit) => hit || fetch(req).then((res) => {
          if (res.ok || res.type === 'opaque') c.put(req, res.clone());
          return res;
        }))
      )
    );
  }
});
