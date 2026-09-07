# Desafio técnico — VB.NET e API

Aplicação de estudo em **Visual Basic .NET 10 / Windows Forms** que consulta uma API HTTP usando **MSXML 6** e exibe nome e preço dos produtos em um **ListView**.

Os produtos são fictícios. A API de exemplo roda em Node.js 24, com cenários de lista vazia, HTTP 503, JSON inválido, preço inválido e timeout.

## Executar

Requisitos: Windows com MSXML 6, Visual Studio 2026 com Desenvolvimento para desktop com .NET e Docker Desktop usando containers Linux.

Na pasta raiz deste repositório, abra o PowerShell e crie o container da API:

```powershell
docker run --name ps-estudo-api -p 127.0.0.1:18081:3000 --mount "type=bind,source=$($PWD.Path)\API,target=/app,readonly" -w /app -d node:24-alpine node server.js
```

Se o container já existir, use `docker start ps-estudo-api`. Confira sua configuração antes de reutilizar um container com esse nome.

1. Abra `http://127.0.0.1:18081/produtos` no navegador e confira os três produtos.
2. Abra `desafio-tec-vbnet/desafio-tec-vbnet.slnx` no Visual Studio.
3. Execute com **F5** e clique em **Carregar produtos**.

Para editar a tela, abra **Desafio.vb → Exibir Designer**.

Para encerrar a API:

```powershell
docker stop ps-estudo-api
```

## Testar respostas

Acrescente à URL um destes parâmetros:

| Parâmetro | Resposta esperada |
| --- | --- |
| Sem parâmetro | Três produtos |
| `?cenario=vazio` | Lista vazia |
| `?cenario=erro` | HTTP 503 |
| `?cenario=json` | JSON inválido |
| `?cenario=preco` | Preço negativo rejeitado |
| `?cenario=lento` | Timeout |

Em falhas de consulta, a lista anterior é mantida com uma mensagem. Uma resposta válida e vazia limpa a lista. A consulta roda fora da thread da interface e o botão fica desabilitado durante a espera.

## Organização

- `API/server.js`: API local com produtos e falhas simuladas.
- `desafio-tec-vbnet/Desafio.vb`: evento do botão e apresentação dos resultados.
- `desafio-tec-vbnet/ProdutosApi.vb`: MSXML, status HTTP e validação de JSON.
- `desafio-tec-vbnet/Produto.vb`: nome e preço.

## Estado e limites

- Compilação do projeto VB.NET conferida; execução com sucesso também relatada durante o estudo.
- A API precisa retornar HTTP 200 e uma lista JSON com `nome` não vazio e `preco` numérico não negativo, representável como Decimal.
- Os limites de tempo são por etapa de rede; não representam um prazo total fixo.
- Esta é uma adaptação em **VB.NET**. A versão **VB6**, a função PostgreSQL/ILIKE e o Windows Service em C# ainda não estão incluídos neste repositório.
