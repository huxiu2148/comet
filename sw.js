// app-shell 快取：完全沒有訊號/網路時（打牌現場很常見），只要之前開過
// 一次，重新整理或重開瀏覽器還是能把畫面叫出來（資料本來就存在
// localStorage，不受這個檔案影響，這裡只負責讓「網頁本身」能顯示出來）。
//
// 一律網路優先（能連網就永遠抓最新版本），只有在真的抓不到（離線）才回頭
// 用快取——這樣才不會擋到 index.html 裡既有的「每 5 分鐘檢查新版本、
// 有更新就強制要求重新整理」機制（那個機制自己用 cache:"no-store" 直接
// fetch，如果這裡用快取優先會讓使用者永遠看不到更新提示）。
const CACHE_NAME = "juewei-xifeng-shell-v1";
const APP_SHELL = [
  "index.html",
  "manifest.json",
  "icons/icon-32.png",
  "icons/icon-180.png",
  "icons/icon-192.png",
  "icons/icon-512.png"
];

self.addEventListener("install", function(event){
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){ return cache.addAll(APP_SHELL); })
  );
});

self.addEventListener("activate", function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k !== CACHE_NAME; }).map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener("fetch", function(event){
  var req = event.request;
  var url = new URL(req.url);
  if(req.method !== "GET" || url.origin !== self.location.origin) return; // 跨網域（Firebase 等）一律不插手

  var fileName = url.pathname.split("/").filter(Boolean).pop() || "index.html";
  if(APP_SHELL.indexOf(fileName) === -1 && url.pathname !== "/" && !url.pathname.endsWith("/")) return;
  var cacheKey = (fileName === "index.html" || url.pathname === "/" || url.pathname.endsWith("/")) ? "index.html" : fileName;

  event.respondWith(
    fetch(req).then(function(res){
      if(res && res.ok){
        var copy = res.clone();
        caches.open(CACHE_NAME).then(function(cache){ cache.put(cacheKey, copy); });
      }
      return res;
    }).catch(function(){
      return caches.match(cacheKey).then(function(cached){ return cached || caches.match("index.html"); });
    })
  );
});
