/* Оболочка сайта отдаётся мгновенно из памяти телефона,
   а новая версия подтягивается в фоне. Данные всегда берутся из сети. */
const CACHE = 'att-v2';
const SHELL = ['./', 'index.html', 'manifest.json',
  'assets/logo-full.png', 'assets/logo-full-inverse.png',
  'assets/logo-mark.png', 'assets/logo-mark-inverse.png', 'assets/icon-192.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if(e.request.method !== 'GET' || url.origin !== location.origin) return;  // запросы к Apps Script не трогаем

  e.respondWith(
    caches.match(e.request).then(hit => {
      const fresh = fetch(e.request).then(res => {
        if(res && res.ok){
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        }
        return res;
      }).catch(() => hit);
      return hit || fresh;          // есть в памяти — отдаём сразу, обновление приедет к следующему открытию
    })
  );
});
