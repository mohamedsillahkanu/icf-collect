// ICF Collect Service Worker for Offline Support
const CACHE_NAME = 'icf-collect-v8';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './manifest.json'
];

// External assets to cache
const EXTERNAL_ASSETS = [
    'https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;700&display=swap',
    'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Install - cache core assets
self.addEventListener('install', event => {
    console.log('Service Worker installing...');
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                console.log('Caching core assets');
                // Cache local assets
                return cache.addAll(ASSETS_TO_CACHE);
            })
            .then(() => {
                // Try to cache external assets (don't fail if some don't work)
                return caches.open(CACHE_NAME).then(cache => {
                    return Promise.allSettled(
                        EXTERNAL_ASSETS.map(url => 
                            fetch(url, { mode: 'cors' })
                                .then(response => {
                                    if (response.ok) {
                                        return cache.put(url, response);
                                    }
                                })
                                .catch(err => console.log('Could not cache:', url))
                        )
                    );
                });
            })
            .then(() => self.skipWaiting())
    );
});

// Activate - clean up old caches
self.addEventListener('activate', event => {
    console.log('Service Worker activating...');
    event.waitUntil(
        caches.keys().then(keys => {
            return Promise.all(
                keys.filter(key => key !== CACHE_NAME)
                    .map(key => caches.delete(key))
            );
        }).then(() => self.clients.claim())
    );
});

// Fetch - network first for API, cache first for assets
self.addEventListener('fetch', event => {
    const url = new URL(event.request.url);
    
    // Skip non-GET requests
    if (event.request.method !== 'GET') {
        return;
    }
    
    // For Google Apps Script API calls - always try network
    if (url.hostname.includes('script.google.com') || 
        url.hostname.includes('script.googleusercontent.com')) {
        event.respondWith(
            fetch(event.request).catch(() => {
                return new Response(JSON.stringify({ success: false, error: 'Offline' }), {
                    headers: { 'Content-Type': 'application/json' }
                });
            })
        );
        return;
    }
    
    // For navigation requests (HTML pages with ?d= params)
    if (event.request.mode === 'navigate') {
        event.respondWith(
            fetch(event.request)
                .then(response => {
                    // Cache the response
                    const responseClone = response.clone();
                    caches.open(CACHE_NAME)
                        .then(cache => cache.put('./index.html', responseClone));
                    return response;
                })
                .catch(() => {
                    // Offline - serve cached index.html
                    console.log('Offline: serving cached page');
                    return caches.match('./index.html');
                })
        );
        return;
    }
    
    // For other requests - cache first, then network
    event.respondWith(
        caches.match(event.request)
            .then(cached => {
                if (cached) {
                    return cached;
                }
                return fetch(event.request)
                    .then(response => {
                        if (response.ok) {
                            const responseClone = response.clone();
                            caches.open(CACHE_NAME)
                                .then(cache => cache.put(event.request, responseClone));
                        }
                        return response;
                    });
            })
    );
});
