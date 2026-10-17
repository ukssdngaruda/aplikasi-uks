// Service worker minimal agar aplikasi bisa dipasang (install) di HP.
// Tidak menyimpan cache, jadi data selalu terbaru dari server.
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () { /* langsung ke jaringan */ });
