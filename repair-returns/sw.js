'use strict';
const prefix='cbtrade-repair-assets-'+new URL(self.registration.scope).pathname+'-';
const cacheName=prefix+'1.0.0-r2';
const assets=['./','./index.html','./app.js','./word.js','./data.js','./app.css','./jszip.min.js','./icon-180.png','./icon-192.png','./icon-512.png','./manifest.webmanifest'];
self.addEventListener('install',e=>e.waitUntil(caches.open(cacheName).then(c=>c.addAll(assets))));
self.addEventListener('activate',e=>e.waitUntil((async()=>{for(const name of await caches.keys())if(name.startsWith(prefix)&&name!==cacheName)await caches.delete(name);await self.clients.claim();})()));
self.addEventListener('fetch',e=>{
 const url=new URL(e.request.url);
 if(e.request.method!=='GET'||!url.href.startsWith(self.registration.scope))return;
 const relative='./'+url.pathname.slice(new URL(self.registration.scope).pathname.length);
 if(!assets.includes(relative))return;
 e.respondWith((async()=>{const c=await caches.open(cacheName);const cached=await c.match(relative);if(cached)return cached;const response=await fetch(e.request);if(response.ok)await c.put(relative,response.clone());return response;})());
});
