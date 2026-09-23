// เก็บไฟล์ไลบรารีจาก CDN (ตัวอ่านเลขไมล์ ~7MB, Excel, Word) ไว้ในเครื่อง
// เปิดแอปครั้งถัดไปจะไม่ต้องดาวน์โหลดใหม่ — ตัวหน้าเว็บเอง (index.html) ยังโหลดจากเน็ตตามปกติเพื่อให้ได้เวอร์ชันล่าสุดเสมอ
const CACHE = 'lib-cache-v1';
const CDN_HOSTS = ['cdn.jsdelivr.net', 'unpkg.com'];

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(k => k.startsWith('lib-cache-') && k !== CACHE).map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  let url;
  try { url = new URL(req.url); } catch (e) { return; }
  if (!CDN_HOSTS.includes(url.hostname)) return; // อย่างอื่น (หน้าแอป, Google Maps) ปล่อยผ่านตามปกติ
  event.respondWith((async () => {
    const cache = await caches.open(CACHE);
    const hit = await cache.match(req, { ignoreVary: true });
    if (hit) return hit;
    const res = await fetch(req);
    if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone()).catch(() => {});
    return res;
  })());
});
