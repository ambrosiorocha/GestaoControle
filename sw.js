const CACHE_NAME = 'Gestão&Controle-v108';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './css/style.css',
    './js/config.js',
    './js/auth.js',
    './js/main.js',
    './js/dashboard.js',
    './js/vendas.js',
    './js/produtos.js',
    './js/clientes.js',
    './js/fornecedores.js',
    './js/financeiro.js',
    './js/relatorios.js',
    './Vendas.html',
    './Produtos.html',
    './Clientes.html',
    './Fornecedores.html',
    './Financeiro.html',
    './Relatorios.html',
    './manifest_v107.json',
    './assets/logo.png',
    './icons/icon-192-v4.png',
    './icons/icon-512-v4.png'
];

// Instalação do Service Worker e cache dos recursos
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                console.log('Cache aberto com sucesso');
                return cache.addAll(ASSETS_TO_CACHE);
            })
    );
    self.skipWaiting();
});

// Interceptar as requisições (Fetch) - ESTRATÉGIA: NETWORK FIRST
self.addEventListener('fetch', event => {
    // Ignorar requisições ao Google Apps Script ou métodos diferentes de GET
    if (event.request.url.includes('script.google.com') || event.request.method !== 'GET') {
        event.respondWith(fetch(event.request));
        return;
    }

    event.respondWith(
        fetch(event.request)
            .then(response => {
                // Se der certo e for uma resposta válida, atualiza o cache (em background)
                if (response && response.status === 200 && response.type === 'basic') {
                    const responseToCache = response.clone();
                    caches.open(CACHE_NAME).then(cache => {
                        cache.put(event.request, responseToCache);
                    });
                }
                return response;
            })
            .catch(() => {
                // Se falhar a rede (ex: offline), pega do cache
                return caches.match(event.request);
            })
    );
});

// Limpeza de caches antigos (Ativação)
self.addEventListener('activate', event => {
    const cacheWhitelist = [CACHE_NAME];
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheWhitelist.indexOf(cacheName) === -1) {
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
    self.clients.claim();
});
