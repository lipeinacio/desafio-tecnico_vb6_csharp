# VB6 — consulta de produtos

Requer VB6, MSXML 6 e Microsoft Windows Common Controls 6.0 (`MSCOMCTL.OCX`).

1. Na raiz do repositório, execute `node 01-vb6/API/server.js` (requer Node.js).
2. Abra `DesafioVB6.vbp` no VB6 e pressione F5.
3. Use `http://127.0.0.1:3000/produtos` e clique em **Carregar produtos**.

Se o VB6 estiver em uma VM com NAT do VirtualBox e a API no host, use `http://10.0.2.2:3000/produtos`. A URL pode ser alterada na tela ou em `url.txt`.

## Código

- `frmDesafio.frm`: layout e fluxo do botão, espera e exibição.
- `CProdutosApi.cls`: inicia a consulta MSXML e entrega a resposta validada.
- `CProduto.cls`: nome e preço de um produto.
- `apoio/`: leitura de JSON, formato de moeda, mensagens e relógio do timeout.
- `testes/`: testes do parser e da API, executados com `DesafioVB6.exe /test` após compilar.

Uma falha preserva a lista anterior; lista vazia limpa os itens. A consulta tem limite total de três segundos. Para testar os erros, acrescente à URL `?cenario=vazio`, `?cenario=erro`, `?cenario=json`, `?cenario=preco` ou `?cenario=lento`.

O projeto completo está no `.vbp`. O layout está no `.frm` e não utiliza recursos `.frx`.
