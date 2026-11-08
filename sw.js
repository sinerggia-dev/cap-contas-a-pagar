/* Service worker minimo: existe para o Chrome do Android reconhecer o CAP como
   app instalavel (abre em tela cheia, sem barra de endereco). Nao guarda nada em
   cache — toda requisicao vai direto para a rede, entao nunca mostra versao velha. */
self.addEventListener('install', function(){ self.skipWaiting(); });
self.addEventListener('activate', function(ev){ ev.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function(){ /* sem respondWith: o navegador busca na rede */ });
