# C# — monitor de arquivos

Requer .NET SDK 10. Os comandos abaixo são executados na raiz do repositório.

Configure entrada e logs em `02-csharp/MonitorArquivos/appsettings.json`. Caminhos relativos usam a pasta do executável como base.

```powershell
dotnet run --project 02-csharp/MonitorArquivos
dotnet run --project 02-csharp/Testes
```

Crie e exclua arquivos na pasta configurada. Os eventos são gravados em `eventos.log`. Ctrl+C encerra o modo console. Os testes usam uma pasta temporária e mostram o caminho do log.

## Código

- `MonitorArquivos/Program.cs`: inicia a aplicação e registra o Windows Service.
- `MonitorArquivos/MonitorWorker.cs`: inicia a observação, trata criação/exclusão, grava o log e encerra o monitor.
- `Testes/Program.cs`: verifica inicialização, criação, exclusão, lote de arquivos e encerramento.

Para instalar como serviço, publique e execute o instalador em PowerShell como administrador:

```powershell
dotnet publish 02-csharp/MonitorArquivos -c Release -o publicado
.\02-csharp\MonitorArquivos\Instalar-Servico.ps1 -PastaPublicacao .\publicado
```

O instalador usa `PsMonitorArquivos` e não substitui um serviço existente. A conta do serviço precisa acessar as pastas configuradas. Somente a pasta indicada é monitorada, sem subpastas.
