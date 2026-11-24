'use strict';
const prefix='cbtrade-repair-assets-'+new URL(self.registration.scope).pathname+'-';
const cacheName=prefix+'1.1.1';
const assets=['./','./mobile-1.1.1.html','./index.html','./app-1.1.1.js','./word-1.1.1.js','./data-1.1.1.js','./app-1.1.1.css','./jszip.min.js','./icon-180.png','./icon-192.png','./icon-512.png','./manifest.webmanifest'];
self.addEventListener('install',e=>e.waitUntil(caches.open(cacheName).then(c=>c.addAll(assets.map(path=>new Request(new URL(path,self.registration.scope),{cache:'reload'})))).then(()=>self.skipWaiting())));
self.addEventListener('activate',e=>e.waitUntil((async()=>{for(const name of await caches.keys())if(name.startsWith(prefix)&&name!==cacheName)await caches.delete(name);await self.clients.claim();})()));
self.addEventListener('fetch',e=>{
 const url=new URL(e.request.url);
 if(e.request.method!=='GET'||!url.href.startsWith(self.registration.scope))return;
 const relative='./'+url.pathname.slice(new URL(self.registration.scope).pathname.length);
 if(!assets.includes(relative))return;
 e.respondWith((async()=>{const c=await caches.open(cacheName);const cached=await c.match(relative);if(cached)return cached;const response=await fetch(e.request);if(response.ok)await c.put(relative,response.clone());return response;})());
});
