const CACHE = 'agf-v3';
const BASE = self.location.pathname.replace(/sw\.js$/, '');
const ASSETS = [
  BASE,
  BASE + 'index.html',
  BASE + 'manifest.json',
  BASE + 'icon.svg'
];
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzi9aawyYOksTthOJ5HQMgSE8ULX06R-ZmcTdBzH03BauaxT5az7gKL_cw9eULjL5Cx/exec";

self.addEventListener('install', e => {
  e.waitUntil(
    Promise.all([
      caches.open(CACHE).then(c => c.addAll(ASSETS)),
      self.skipWaiting()
    ])
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // Pass cross-origin requests through (Apps Script, dolar API, etc.)
  if (url.origin !== location.origin) return;
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(resp => {
        const clone = resp.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return resp;
      }).catch(() => caches.match(BASE + 'index.html'));
    })
  );
});

// Background Sync
self.addEventListener('sync', e => {
  if (e.tag === 'agf-sync') e.waitUntil(bgSync());
});

async function bgSync() {
  const db = await idbOpen();
  const all = await idbGetAll(db);
  for (const item of all) {
    await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      body: JSON.stringify(item.payload)
    });
    await idbDelete(db, item.id);
  }
}

function idbOpen() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('agf_db', 1);
    r.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('pending_ventas'))
        db.createObjectStore('pending_ventas', { keyPath: 'id', autoIncrement: true });
    };
    r.onsuccess = e => res(e.target.result);
    r.onerror = e => rej(e.target.error);
  });
}
function idbGetAll(db) {
  return new Promise((res, rej) => {
    const r = db.transaction('pending_ventas', 'readonly').objectStore('pending_ventas').getAll();
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
function idbDelete(db, id) {
  return new Promise((res, rej) => {
    const r = db.transaction('pending_ventas', 'readwrite').objectStore('pending_ventas').delete(id);
    r.onsuccess = () => res();
    r.onerror = () => rej(r.error);
  });
}
