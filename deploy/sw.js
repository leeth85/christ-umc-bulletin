/* Christ UMC Digital Bulletin Service Worker (fast-open, offline-capable). Keep this URL stable. */
const ROOT = new URL("./", self.registration.scope);
const HOME = new URL("index.html", ROOT).href;
const CACHE = "cumc-bulletin-recovery-v1-" + encodeURIComponent(ROOT.pathname);
const OWNED = /^(cumc-bulletin-shell-|dasibon-bulletin-shell-|dasibon-pwa-engine-)/;
const ASSETS = ["manifest.json", "icon-192.png", "icon-512.png"].map((p) => new URL(p, ROOT).href);
const EXTERNAL = new Set(["fonts.googleapis.com", "fonts.gstatic.com", "cdnjs.cloudflare.com"]);
const BUILD = "20261123-cumc-1";
const TIMEOUT = 5000;
const FETCH_LIMIT = 90000;
const RECENT = 60000;
let lastFresh = 0, lastVersion = "", cachedVersion = null;

function isHome(url) {
  return url.origin === ROOT.origin && [ROOT.pathname, new URL(HOME).pathname].includes(url.pathname);
}
function version(text) {
  const release = /data-bulletin-release=["']([^"']+)/.exec(text);
  const build = /data-pwa-build=["']([^"']+)/.exec(text);
  return (release ? release[1] : "") + ":" + (build ? build[1] : "legacy");
}
function navigationResponse(response) {
  const headers = new Headers(response.headers);
  headers.delete("content-encoding");
  headers.delete("content-length");
  headers.set("Cache-Control", "no-store");
  return new Response(response.body, { status: response.status, statusText: response.statusText, headers });
}
function validText(text) {
  return /id=["']dasibonApp["']/.test(text) && /id=["']panel-home["']/.test(text) && text.includes("/*STORE_END*/") && /<\/html>/i.test(text);
}
async function valid(response) {
  if (!response || !response.ok || !/text\/html/i.test(response.headers.get("content-type") || "")) return false;
  return validText(await response.clone().text());
}
async function read(cacheName, key) {
  try { return await (await caches.open(cacheName)).match(key); } catch (_) { return undefined; }
}
async function save(key, response) {
  try { await (await caches.open(CACHE)).put(key, response.clone()); return true; } catch (_) { return false; }
}
async function savedCopy() {
  const current = await read(CACHE, HOME);
  if (!current) return null;
  const text = await current.clone().text();
  if (!validText(text)) return null;
  cachedVersion = version(text.slice(0, 4096));
  return current;
}
async function fallback() {
  const current = await savedCopy();
  if (current) return navigationResponse(current);
  let names = [];
  try { names = await caches.keys(); } catch (_) {}
  for (const name of names.filter((n) => OWNED.test(n)).reverse()) {
    for (const key of [HOME, ROOT.href]) {
      const old = await read(name, key);
      if (await valid(old)) {
        const safe = navigationResponse(old);
        await save(HOME, safe);
        cachedVersion = null;
        return safe;
      }
    }
  }
  return null;
}
async function cleanup() {
  if (!await savedCopy()) return;
  try {
    for (const name of (await caches.keys()).filter((n) => OWNED.test(n))) {
      const cache = await caches.open(name), keys = await cache.keys();
      if (keys.every((k) => {
        const u = new URL(k.url);
        return EXTERNAL.has(u.hostname) || (u.origin === ROOT.origin && u.pathname.startsWith(ROOT.pathname));
      })) await caches.delete(name);
    }
  } catch (_) {}
}
async function inform(message, clientId) {
  try {
    const clients = clientId ? [await self.clients.get(clientId)] : await self.clients.matchAll({ type: "window" });
    clients.filter(Boolean).filter((c) => isHome(new URL(c.url))).forEach((c) => c.postMessage(message));
  } catch (_) {}
}
async function fresh() {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FETCH_LIMIT);
  try {
    if (cachedVersion === null) await savedCopy();
    const url = new URL(HOME);
    url.searchParams.set("_bulletin_fetch", Date.now().toString());
    const response = await fetch(url.href, { cache: "no-store", signal: controller.signal, credentials: "same-origin" });
    if (response.redirected && !isHome(new URL(response.url))) throw new Error("Unexpected redirect");
    if (!response.ok || !/text\/html/i.test(response.headers.get("content-type") || "") || !response.body) throw new Error("Invalid bulletin response");
    const reader = response.body.getReader(), decoder = new TextDecoder(), chunks = [];
    let head = "";
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      if (head.length < 4096) head += decoder.decode(value, { stream: true });
      if (head.length >= 4096 || /<head[\s>]/i.test(head)) break;
    }
    const latest = version(head);
    if (cachedVersion && latest === cachedVersion) {
      reader.cancel().catch(() => {});
      controller.abort();
      lastFresh = Date.now();
      lastVersion = latest;
      await inform({ type: "BULLETIN_FRESH", version: latest });
      return null;
    }
    await inform({ type: "BULLETIN_DOWNLOADING" });
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const full = new Response(new Blob(chunks), { status: response.status, statusText: response.statusText, headers: response.headers });
    if (!await valid(full)) throw new Error("Invalid or incomplete bulletin");
    const safe = navigationResponse(full);
    await save(HOME, safe);
    cachedVersion = latest;
    lastFresh = Date.now();
    lastVersion = latest;
    await cleanup();
    await inform({ type: "BULLETIN_FRESH", version: latest });
    return safe;
  } finally {
    clearTimeout(timer);
  }
}
let pending;
function refresh() {
  if (!pending) pending = fresh().finally(() => { pending = null; });
  return pending;
}
function bounded(promise, ms = TIMEOUT) {
  let timer;
  return Promise.race([promise, new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error("Timed out")), ms);
  })]).finally(() => clearTimeout(timer));
}
function guidance() {
  return new Response(
    '<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Christ UMC Digital Bulletin · Offline</title>' +
    '<style>body{font:17px/1.65 -apple-system,sans-serif;background:#F6F1E8;color:#1E1F4B;margin:10vh auto;padding:24px;max-width:520px}a{display:inline-block;padding:12px 20px;background:#1E1F4B;color:#fff;border-radius:10px;text-decoration:none;font-weight:700}</style>' +
    '<h1>Unable to Load Bulletin</h1>' +
    '<p>Please check your internet connection. Once you open the bulletin online once, it will remain available offline during worship.</p>' +
    '<a href="./index.html">Try Again</a>' +
    "<script>addEventListener('online',()=>location.reload());if('serviceWorker'in navigator)navigator.serviceWorker.register('./sw.js',{scope:'./',updateViaCache:'none'}).then(r=>r.update()).catch(()=>{});<\/script></html>",
    { status: 503, headers: { "Content-Type": "text/html;charset=utf-8", "Cache-Control": "no-store" } }
  );
}
self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    try { await bounded(refresh(), 15000); } catch (_) { await fallback(); }
    await self.skipWaiting();
  })());
});
self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    await fallback();
    await cleanup();
    if (self.registration.navigationPreload) await self.registration.navigationPreload.disable().catch(() => {});
    await self.clients.claim();
  })());
});
self.addEventListener("message", (event) => {
  const d = event.data || {};
  if (d.type === "GET_BULLETIN_BUILD" && event.ports && event.ports[0]) event.ports[0].postMessage({ build: BUILD });
  if (d.type === "SKIP_WAITING") event.waitUntil(self.skipWaiting());
  if (d.type === "CHECK_BULLETIN") event.waitUntil((async () => {
    const cid = event.source ? event.source.id : undefined;
    try {
      if (pending) await pending;
      else if (Date.now() - lastFresh < RECENT) await inform({ type: "BULLETIN_FRESH", version: lastVersion }, cid);
      else await refresh();
    } catch (_) {
      await inform({ type: "BULLETIN_OFFLINE" }, cid);
    }
  })());
});
self.addEventListener("fetch", (event) => {
  const request = event.request, url = new URL(request.url);
  if (request.method !== "GET") return;
  if (request.mode === "navigate" && isHome(url)) {
    const savedP = fallback().catch(() => null);
    event.waitUntil(savedP.then((s) => s && Date.now() - lastFresh >= RECENT ? refresh().catch(() => {}) : null));
    event.respondWith((async () => {
      const saved = await savedP;
      if (saved) return saved;
      try {
        const r = await refresh();
        if (r) return r.clone();
        return await fallback() || guidance();
      } catch (_) {
        return await fallback() || guidance();
      }
    })());
    return;
  }
  if (!ASSETS.includes(url.href) && !EXTERNAL.has(url.hostname)) return;
  const network = fetch(request, { cache: "no-cache" }).then(async (r) => {
    if (r.ok || r.type === "opaque") await save(request, r);
    return r;
  });
  event.waitUntil(network.catch(() => {}));
  event.respondWith((async () => {
    try {
      const r = await bounded(network);
      if (r.ok || r.type === "opaque") return r;
    } catch (_) {}
    return await read(CACHE, request) || new Response("", { status: 503 });
  })());
});
