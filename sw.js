/* PharmEase service worker — keeps the app itself (and the Firebase code it needs) available offline.
   Your data is kept separately by Firebase, on the device.
   To publish an update, change CACHE to a new name, e.g. 'pharmease-v4'. */
const CACHE = 'pharmease-v6';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png'];
const SDK = ['app', 'auth', 'firestore'].map(n => 'https://www.gstatic.com/firebasejs/10.14.1/firebase-' + n + '.js');
const ALLOWED = ['www.gstatic.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

// Save a Firebase module and every module it imports, so the app can start with no internet.
async function cacheModule(cache, url, seen) {
  if (seen.has(url)) return; seen.add(url);
  const res = await fetch(url); if (!res.ok) return;
  await cache.put(url, res.clone());
  const text = await res.text(), re = /(?:from|import)\s*\(?\s*["'](\.{1,2}\/[^"']+)["']/g, next = [];
  let m; while ((m = re.exec(text))) next.push(new URL(m[1], url).href);
  await Promise.all(next.map(u => cacheModule(cache, u, seen)));
}

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(async c => {
    await c.addAll(SHELL);
    const seen = new Set();
    await Promise.all(SDK.map(u => cacheModule(c, u, seen).catch(() => {})));
  }).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Only handle our own files, the pinned Firebase code and fonts. Everything else
  // (Firestore and sign-in traffic) goes straight to the network, untouched.
  if (url.origin !== location.origin && !ALLOWED.includes(url.hostname)) return;

  if (req.mode === 'navigate') {            // opening the app: newest copy if online, saved copy if not
    e.respondWith(fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put('./index.html', copy));
      return res;
    }).catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(req).then(hit => {      // saved copy first, refreshed in the background
    const net = fetch(req).then(res => {
      if (res.ok || res.type === 'opaque') { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => hit);
    return hit || net;
  }));
});
