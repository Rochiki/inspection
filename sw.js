/* =====================================================================
   Service worker — membuat aplikasi tetap jalan tanpa internet.
   SETIAP KALI ADA PERUBAHAN FILE: naikkan angka VERSI di bawah (v1 -> v2 -> v3 ...)
   supaya HP dan laptop pengguna mengambil versi terbaru.
   ===================================================================== */
const VERSI = "v1";
const CACHE = "temuan-qshe-" + VERSI;
const FILES = [
  "./", "./index.html", "./manifest.webmanifest", "./icon-192.png", "./icon-512.png",
  "./css/style.css", "./css/fonts.css",
  "./js/data-klausul.js", "./js/data-ppt.js", "./js/rapikan.js", "./js/app.js",
  "./lib/exceljs.min.js", "./lib/pptxgen.bundle.js", "./lib/chart.umd.js",
  "./fonts/barlow-latin-400-normal.woff2", "./fonts/barlow-latin-500-normal.woff2",
  "./fonts/barlow-latin-600-normal.woff2", "./fonts/barlow-latin-700-normal.woff2",
  "./fonts/barlow-condensed-latin-600-normal.woff2", "./fonts/barlow-condensed-latin-700-normal.woff2"
];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("temuan-qshe-") && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  if (e.request.method !== "GET" || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(caches.open(CACHE).then(c => c.match(e.request, { ignoreSearch: true }).then(hit => {
    const net = fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => null);
    return hit || net.then(r => r || c.match("./index.html"));
  })));
});
