/* Страница всегда берётся из сети — так обновления доезжают сразу.
   Картинки и манифест отдаём из памяти. Данные (Apps Script) не трогаем вовсе.
   Мгновенное открытие обеспечивает не этот кеш, а сохранённые данные в localStorage. */
const CACHE = 'att-v3';
const ASSETS = ['manifest.json',
  'assets/logo-full.png', 'assets/logo-full-inverse.png',
  'assets/logo-mark.png', 'assets/logo-mark-inverse.png',
  'assets/icon-192.png', 'assets/icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

const isPage = req => req.mode === 'navigate' || req.destination === 'document' ||
                      /\/(index\.html)?(\?.*)?$/.test(new URL(req.url).pathname + new URL(req.url).search);

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if(e.request.method !== 'GET' || url.origin !== location.origin) return;

  if(isPage(e.request)){
    // страница: сначала сеть, кеш — только если связи нет
    e.respondWith(
      fetch(e.request, { cache: 'no-store' })
        .then(res => {
          if(res && res.ok){
            const copy = res.clone();
            caches.open(CACHE).then(c => c.put('index.html', copy)).catch(() => {});
          }
          return res;
        })
        .catch(() => caches.match('index.html').then(r => r || Response.error()))
    );
    return;
  }

  // остальное: отдаём из памяти, в фоне обновляем
  e.respondWith(
    caches.match(e.request).then(hit => {
      const fresh = fetch(e.request).then(res => {
        if(res && res.ok){
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {});
        }
        return res;
      }).catch(() => hit);
      return hit || fresh;
    })
  );
});
