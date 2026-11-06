// تقاريري — Service Worker: يجعل التطبيق يفتح ويعمل دون إنترنت
const V = 'taqarir-v1';
const SHELL = [
  './', './index.html', './manifest.json', './icon-192.png', './icon-512.png',
  './lib/html-to-image.js', './lib/jspdf.umd.min.js',
  './fonts/tajawal-arabic-400-normal.woff2', './fonts/tajawal-arabic-500-normal.woff2',
  './fonts/tajawal-arabic-700-normal.woff2', './fonts/tajawal-arabic-800-normal.woff2',
  './fonts/tajawal-latin-400-normal.woff2', './fonts/tajawal-latin-500-normal.woff2',
  './fonts/tajawal-latin-700-normal.woff2', './fonts/tajawal-latin-800-normal.woff2'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== V).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // طلبات الخادم تمر مباشرة

  // الصفحة: الشبكة أولاً (للحصول على آخر تحديث)، ثم النسخة المحفوظة
  if (req.mode === 'navigate' || url.pathname.endsWith('/index.html')) {
    e.respondWith(
      fetch(req).then(r => { const cp = r.clone(); caches.open(V).then(c => c.put('./index.html', cp)); return r; })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }
  // باقي الملفات: المحفوظ أولاً
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (r.ok) { const cp = r.clone(); caches.open(V).then(c => c.put(req, cp)); }
    return r;
  })));
});
