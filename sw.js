// オンライン用。アプリ本体を古いキャッシュに固定しない。
self.addEventListener('install', (event) => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  const scope = new URL(self.registration.scope);
  const isAppFile = event.request.mode === 'navigate' ||
    /\.(html|js|css|json|webmanifest|png|ico|svg)$/.test(url.pathname);

  // 外部の GAS 通信や書き込みリクエストには介入しない。
  if (event.request.method !== 'GET' || url.origin !== scope.origin ||
      !url.pathname.startsWith(scope.pathname) || !isAppFile) return;

  event.respondWith(fetch(new Request(event.request, { cache: 'no-store' })));
});
