// 시골 펜션 앱 껍데기 — 서비스 워커
//  · 껍데기 파일(아이콘 · manifest)만 저장해 둠
//  · index.html 은 항상 새로 받음 (인터넷이 없을 때만 저장본)
//  · 구글 웹앱(다른 주소) 요청은 건드리지 않음
var CACHE = "pension-shell-v1";
var FILES = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png",
  "./icon-maskable-512.png", "./apple-touch-icon.png", "./favicon.png"];

self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(FILES); }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;   // 구글 웹앱은 그대로
  if (req.mode === "navigate") {
    // 화면(index.html): 먼저 새로 받고, 안 되면 저장본 (주소 뒤 ?k= 는 무시하고 찾음)
    e.respondWith(fetch(req).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put("./index.html", copy); });
      return res;
    }).catch(function () { return caches.match("./index.html"); }));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(function (hit) { return hit || fetch(req); }));
});
