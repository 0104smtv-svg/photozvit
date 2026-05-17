const CACHE_NAME = 'fotozvit-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// Встановлення — кешуємо файли
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

// Активація — видаляємо старий кеш
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys
        .filter(key => key !== CACHE_NAME)
        .map(key => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// Запити — спочатку мережа, fallback на кеш
// Запити до Apps Script завжди йдуть через мережу (не кешуються)
self.addEventListener('fetch', e => {
  const url = e.request.url;

  // Запити до Google Script — тільки мережа
  if (url.includes('script.google.com') || url.includes('googleapis.com')) {
    e.respondWith(fetch(e.request).catch(() => new Response(
      JSON.stringify({ status: 'error', message: 'Немає з\'єднання' }),
      { headers: { 'Content-Type': 'application/json' } }
    )));
    return;
  }

  // Решта — мережа з fallback на кеш
  e.respondWith(
    fetch(e.request)
      .then(res => {
        // Оновлюємо кеш свіжою версією
        const clone = res.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
