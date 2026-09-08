const http = require('node:http');

const produtos = [
  { nome: 'Caderno', preco: 24.90 },
  { nome: 'Caneta', preco: 3.50 },
  { nome: 'Mochila', preco: 129.99 }
];

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  const enviar = (status, body) => {
    res.writeHead(status, {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    });
    res.end(typeof body === 'string' ? body : JSON.stringify(body));
  };

  if (req.method !== 'GET') return enviar(405, { erro: 'Use GET' });
  if (url.pathname !== '/produtos') return enviar(404, { erro: 'Rota inexistente' });

  switch (url.searchParams.get('cenario')) {
    case 'vazio':
      return enviar(200, []);
    case 'erro':
      return enviar(503, { erro: 'Falha simulada' });
    case 'json':
      return enviar(200, '{ JSON quebrado');
    case 'preco':
      return enviar(200, [{ nome: 'Teste', preco: -1 }]);
    case 'lento': {
      const timer = setTimeout(() => enviar(200, produtos), 8000);
      res.on('close', () => clearTimeout(timer));
      return;
    }
    default:
      return enviar(200, produtos);
  }
}).listen(3000, '0.0.0.0', () => {
  console.log('API de estudo iniciada na porta 3000');
});