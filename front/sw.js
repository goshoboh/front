// sw.js
self.addEventListener("install", (event) => {
  // 必要ならここでキャッシュとか
  console.log("Service Worker: install");
});

self.addEventListener("activate", (event) => {
  console.log("Service Worker: activate");
});

self.addEventListener("fetch", (event) => {
  // ここでキャッシュ制御すればオフライン対応もできる
  // 今は素通り
});
