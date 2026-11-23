// 영주 할매 — 서비스 워커
//  · 앱 파일(index.html · 아이콘 · manifest · 세방고딕 2.0): 저장본을 바로 보여 주고, 뒤에서 새 버전을 받아 둠 (다음 실행부터 반영)
//  · 글꼴(구글 글꼴 · jsdelivr): 한 번 받으면 저장해 두고 씀
//  · 서버(script.google.com) 요청은 저장하지 않음
//  · ★ 앱을 고칠 때마다 아래 CACHE 숫자를 하나 올려 주세요 (v4 → v5 …)
//      → 쓰는 사람 폰이 새 버전을 통째로 다시 받고, 화면 위에 "새 버전이 준비됐어요" 알림이 뜸
var CACHE = "pension-app-v5", FONTS = "pension-fonts-v1";
var FILES = ["./", "./index.html", "./manifest.json", "./icon-192.png", "./icon-512.png",
  "./icon-maskable-512.png", "./apple-touch-icon.png", "./favicon.png",
  "./SebangGothic2-Regular.woff2", "./SebangGothic2-Bold.woff2"];
var FONT_HOSTS = ["fonts.googleapis.com", "fonts.gstatic.com", "cdn.jsdelivr.net"];

self.addEventListener("install", function (e) {
  // cache: "reload" → 브라우저에 남아 있는 옛날 파일이 아니라 GitHub 에서 새로 받음
  e.waitUntil(caches.open(CACHE).then(function (c) {
    return c.addAll(FILES.map(function (f) { return new Request(f, { cache: "reload" }); }));
  }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE && k !== FONTS; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener("fetch", function (e) {
  var req = e.request, url = new URL(req.url);
  if (req.method !== "GET") return;                                     // 서버 부르기(POST)는 그대로
  if (FONT_HOSTS.indexOf(url.hostname) >= 0) {                          // 글꼴: 저장본 먼저
    e.respondWith(caches.open(FONTS).then(function (c) {
      return c.match(req).then(function (hit) {
        return hit || fetch(req).then(function (res) { if (res.ok || res.type === "opaque") c.put(req, res.clone()); return res; });
      });
    }));
    return;
  }
  if (url.origin !== self.location.origin) return;                      // 구글 서버 등은 건드리지 않음
  var key = req.mode === "navigate" ? "./index.html" : req;
  e.respondWith(caches.open(CACHE).then(function (c) {
    return c.match(key, { ignoreSearch: true }).then(function (hit) {
      var net = fetch(req).then(function (res) { if (res.ok) c.put(key, res.clone()); return res; });
      if (hit) { e.waitUntil(net.catch(function () {})); return hit; }   // 저장본 바로 + 뒤에서 새로 받기
      return net;
    });
  }));
});
