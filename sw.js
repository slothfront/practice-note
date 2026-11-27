// PRACTICE NOTE：ネットが無くても開けるようにする仕組み（サービスワーカー）
// - ページ本体：前回の控えがあればすぐに出し、最新は裏で取りに行く（控えが無いときだけネットを待つ）
// - アイコン・manifest：控えを先に出す
// - 字（Google Fonts）：一度読めたものは控えておく
// 記録（localStorage）と、GASへの送受信には触らない。
const CACHE = 'pnote-v2';
const PAGE = './';
const PRECACHE = ['./', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'apple-touch-icon.png', 'favicon-32.png', 'bg-ball-wine.png', 'bg-ball-navy.png', 'plate-wine.png', 'plate-navy.png'];

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

// 最新のページを取りに行き、取れたら控えを入れ替える（時間を区切らない＝電波が弱くても最後まで取る）
function pageFromNetwork(req) {
  return fetch(req, { cache: 'no-cache' }).then((res) => {
    if (!res.ok) return res;
    const copy = res.clone();
    return caches.open(CACHE).then((c) => c.put(PAGE, copy)).then(() => res, () => res);
  });
}

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // ページ本体：前に開いた控えがあれば、それをすぐに出す（電波が弱くても真っ白で待たない）。
  // 新しい版は裏で取りに行って控えを入れ替える。入れ替わったことは index.html の「新しい版を見つけて入れ替える」が気づいて読み込み直す
  if (req.mode === 'navigate' && url.origin === location.origin) {
    const fresh = pageFromNetwork(req);
    e.waitUntil(fresh.catch(() => {}));
    e.respondWith(
      caches.match(PAGE).then((hit) => hit || fresh).catch(() =>
        caches.match(req, { ignoreSearch: true })
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
