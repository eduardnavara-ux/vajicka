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
  var tasks = [ self.registration.showNotification(title, options) ];
  // Badge na ikoně: u admin objednávek zvedni počítadlo (pokud appka/iOS podporuje)
  if (data.tag === 'objednavka' && 'setAppBadge' in self.navigator) {
    tasks.push(
      (self.__pendingCount = (self.__pendingCount || 0) + 1,
       self.navigator.setAppBadge(self.__pendingCount).catch(function(){}))
    );
  }
  event.waitUntil(Promise.all(tasks));
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

// Appka pošle SW přesný počet (po načtení / odbavení objednávky)
self.addEventListener('message', function(event) {
  var d = event.data || {};
  if (d.type === 'setBadge' && 'setAppBadge' in self.navigator) {
    self.__pendingCount = d.count || 0;
    if (d.count > 0) self.navigator.setAppBadge(d.count).catch(function(){});
    else if ('clearAppBadge' in self.navigator) self.navigator.clearAppBadge().catch(function(){});
  }
});
