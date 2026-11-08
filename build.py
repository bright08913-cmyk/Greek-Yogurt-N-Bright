"""Build the GitHub Pages version of the order page into docs/.
   python3 build.py            (reads the Apps Script /exec link from api_url.txt)"""
import json, pathlib, re
root = pathlib.Path(__file__).parent
SITE = 'https://bright08913-cmyk.github.io/Greek-Yogurt-N-Bright/'
api = (root / 'api_url.txt').read_text().strip()
assert re.match(r'^https://script\.google\.com/macros/s/[\w-]+/exec$', api), 'api_url.txt must be the /exec link'

import hashlib
ver = hashlib.sha1((root / "Index.html").read_bytes()).hexdigest()[:10]
html = (root / 'Index.html').read_text()
head = f'''<link rel="manifest" href="manifest.webmanifest">
<meta name="theme-color" content="#FFF6F1">
<link rel="icon" type="image/png" href="icon-192.png?v=16bit">
<link rel="apple-touch-icon" href="apple-touch-icon.png?v=16bit">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Bright Greek">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="description" content="สั่ง Greek Yogurt — แตะเมนู เลือกผลไม้ แล้วกดสั่งได้เลย">
<meta property="og:type" content="website">
<meta property="og:url" content="{SITE}">
<meta property="og:title" content="Bright Greek — สั่ง Greek Yogurt">
<meta property="og:description" content="แตะเมนู เลือกผลไม้ แล้วกดสั่งได้เลย">
<meta property="og:image" content="{SITE}og-16bit.png">
<meta property="og:image:width" content="1200"><meta property="og:image:height" content="630">
<meta name="twitter:card" content="summary_large_image">
<script>window.GREEK_API = {json.dumps(api)}; window.APP_VER = {json.dumps(ver)};</script>
'''
html = html.replace('<base target="_top">\n', '')          # only needed inside Apps Script
html = html.replace('</title>', '</title>\n' + head, 1)
(root / 'docs' / 'index.html').write_text(html)

(root / 'docs' / 'manifest.webmanifest').write_text(json.dumps({
  "name": "Bright Greek — สั่ง Greek Yogurt", "short_name": "Bright Greek",
  "start_url": "./", "scope": "./", "display": "standalone",
  "background_color": "#FFF4EF", "theme_color": "#FFF4EF", "lang": "th",
  "icons": [
    {"src": "icon-192.png?v=16bit", "sizes": "192x192", "type": "image/png"},
    {"src": "icon-512.png?v=16bit", "sizes": "512x512", "type": "image/png"},
    {"src": "icon-maskable-512.png?v=16bit", "sizes": "512x512", "type": "image/png", "purpose": "maskable"}
  ]}, ensure_ascii=False, indent=2))

# network first (always the newest menu/page), cache only as an offline fallback
(root / 'docs' / 'sw.js').write_text('''const CACHE = 'bright-greek-v3';
self.addEventListener('install', e => { self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', 'index.html', 'manifest.webmanifest', 'icon-192.png']))); });
self.addEventListener('activate', e => e.waitUntil(caches.keys()
  .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
  .then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request, { cache: 'no-cache' }).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put(e.request, c)); return r; })
    .catch(() => caches.match(e.request)));
});
''')
(root / 'docs' / '.nojekyll').write_text('')
(root / 'docs' / 'version.txt').write_text(ver)
print('built docs/ with API', api)
