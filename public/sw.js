const CACHE = "ventadirecta-v1";
const STATIC = ["/", "/index.html"];

self.addEventListener("install", function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(c) { return c.addAll(STATIC); })
  );
  self.skipWaiting();
});

self.addEventListener("activate", function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.filter(function(k){ return k !== CACHE; }).map(function(k){ return caches.delete(k); }));
    })
  );
  self.clients.claim();
});

self.addEventListener("fetch", function(e) {
  // Solo cachear GET, no las llamadas a Supabase
  if (e.request.method !== "GET") return;
  if (e.request.url.includes("supabase.co")) return;
  
  e.respondWith(
    fetch(e.request)
      .then(function(res) {
        // Cachear recursos estáticos (JS, CSS, fonts)
        if (res.ok && (e.request.url.includes("/assets/") || e.request.url.endsWith(".js") || e.request.url.endsWith(".css"))) {
          var clone = res.clone();
          caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        }
        return res;
      })
      .catch(function() {
        // Si no hay red, servir del cache
        return caches.match(e.request).then(function(cached){
          return cached || caches.match("/index.html");
        });
      })
  );
});
