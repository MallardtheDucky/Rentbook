const C = 'rentbook-v4', F = ['./', 'index.html', 'config.js', 'push-config.js', 'manifest.webmanifest', 'icon-192.png'];
self.addEventListener('install', e => e.waitUntil(caches.open(C).then(c => Promise.all(F.map(u => c.add(u).catch(() => {}))))));
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== C).map(x => caches.delete(x))))));
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => { const x = r.clone(); caches.open(C).then(c => c.put(e.request, x)); return r; }).catch(() => caches.match(e.request)));
});

// Phone reminders: the daily message sent from the Google Sheet arrives here.
self.addEventListener('push', e => {
  let j = {};
  try { j = e.data ? e.data.json() : {}; } catch (x) { j = {data: {body: e.data ? e.data.text() : ''}}; }
  const d = Object.assign({}, j.notification || {}, j.data || {});
  e.waitUntil(self.registration.showNotification(d.title || 'Rentbook', {
    body: d.body || '', icon: 'icon-192.png', tag: d.tag || 'rentbook', renotify: true
  }));
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(clients.matchAll({type: 'window', includeUncontrolled: true}).then(list => {
    for (const c of list) if ('focus' in c) return c.focus();
    return clients.openWindow('./');
  }));
});
