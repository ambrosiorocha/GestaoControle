const CACHE_NAME = 'sisvendas-v100';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './Vendas.html',
    './Clientes.html',
    './Fornecedores.html',
    './Produtos.html',
    './Relatorios.html',
    './Financeiro.html',
    './css/style.css',
    './js/main.js',
    './js/vendas.js',
    './js/clientes.js',
    './js/fornecedores.js',
    './js/produtos.js',
    './js/relatorios.js',
    './js/financeiro.js',
    './js/dashboard.js',
    './js/auth.js',
    './js/config.js'
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

// Interceptar as requisições (Fetch)
self.addEventListener('fetch', event => {
    // Ignorar requisições ao Google Apps Script (sempre devem ir para rede)
    if (event.request.url.includes('script.google.com')) {
        event.respondWith(fetch(event.request));
        return;
    }

    event.respondWith(
        caches.match(event.request)
            .then(response => {
                // Retorna do cache se encontrar
                if (response) {
                    return response;
                }
                // Faz a requisição na rede caso contrário
                return fetch(event.request).then(
                    function (response) {
                        // Verifica se a resposta foi válida
                        if (!response || response.status !== 200 || response.type !== 'basic') {
                            return response;
                        }

                        // Clona a resposta e coloca em cache
                        var responseToCache = response.clone();
                        caches.open(CACHE_NAME)
                            .then(function (cache) {
                                cache.put(event.request, responseToCache);
                            });

                        return response;
                    }
                );
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
