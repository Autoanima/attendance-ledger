const CACHE='attendance-ledger-v4';
const APP=['./','./index.html','./manifest.webmanifest','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(APP)).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
// 網頁本體（index.html／導覽用的請求）一律「先試著連網路抓最新版」，
// 抓得到就直接用最新版並更新快取；只有離線抓不到網路時才退回快取的舊版本。
// 這樣以後老師更新這支 App，只要重新整理頁面就能拿到最新版，不會因為舊的離線快取而卡住看不到更新。
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET') return;
  const isPage = e.request.mode==='navigate' || e.request.url.endsWith('/index.html') || e.request.url.endsWith('/');
  if(isPage){
    e.respondWith(
      fetch(e.request).then(r=>{
        const copy=r.clone();
        caches.open(CACHE).then(c=>c.put(e.request,copy));
        return r;
      }).catch(()=>caches.match(e.request))
    );
    return;
  }
  // 其他靜態資源（圖示、manifest 等）維持原本的「快取優先」，減少不必要的重新下載。
  e.respondWith(caches.match(e.request).then(cached=>cached || fetch(e.request).then(r=>{
    const copy=r.clone();
    caches.open(CACHE).then(c=>c.put(e.request,copy));
    return r;
  }).catch(()=>cached)));
});
