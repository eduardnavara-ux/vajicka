/* Service Worker – Dvůr Pod Dubem (web push) */
self.addEventListener('push', function(event) {
  var data = {};
  try { data = event.data ? event.data.json() : {}; } catch(e) {
    data = { title: 'Dvůr Pod Dubem', body: event.data ? event.data.text() : '' };
  }
  var title = data.title || 'Dvůr Pod Dubem';
  var options = {
    body: data.body || '',
    icon: '/leaf-logo.png',
    badge: '/leaf-logo.png',
    tag: data.tag || 'dpd',
    renotify: true,
    data: { url: data.url || '/zakaznik' }
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  var url = (event.notification.data && event.notification.data.url) || '/zakaznik';
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function(list) {
      for (var i = 0; i < list.length; i++) {
        if (list[i].url.indexOf(url) >= 0 && 'focus' in list[i]) return list[i].focus();
      }
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});
