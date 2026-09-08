using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;

var pastaTeste = Path.Combine(Path.GetTempPath(), "MonitorArquivos-Testes", Guid.NewGuid().ToString("N"));
var entrada = Path.Combine(pastaTeste, "entrada");
var logs = Path.Combine(pastaTeste, "logs");
var caminhoLog = Path.Combine(logs, "eventos.log");
var configuracao = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
{
    ["Monitoramento:PastaEntrada"] = entrada,
    ["Monitoramento:PastaLogs"] = logs
}).Build();
using var worker = new MonitorWorker(NullLogger<MonitorWorker>.Instance, configuracao);
var falhas = 0;

string LerLog()
{
    if (!File.Exists(caminhoLog)) return "";
    using var stream = new FileStream(caminhoLog, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    using var reader = new StreamReader(stream);
    return reader.ReadToEnd();
}

async Task Aguardar(Func<bool> condicao)
{
    var inicio = System.Diagnostics.Stopwatch.StartNew();
    while (!condicao())
    {
        if (inicio.Elapsed > TimeSpan.FromSeconds(8))
            throw new Exception("Evento esperado nao foi registrado em 8 segundos.");
        await Task.Delay(25);
    }
}

async Task Testar(string nome, Func<Task> teste)
{
    try { await teste(); Console.WriteLine($"PASS | {nome}"); }
    catch (Exception ex) { falhas++; Console.WriteLine($"FAIL | {nome}: {ex.Message}"); }
}

try
{
    await worker.StartAsync(CancellationToken.None);
    await Testar("Usa as pastas configuradas e registra inicio", async () =>
    {
        await Aguardar(() => LerLog().Contains("INICIADO | " + entrada));
        if (!Directory.Exists(entrada)) throw new Exception("Pasta de entrada nao criada.");
    });

    var arquivo = Path.Combine(entrada, "arquivo-teste.txt");
    await Testar("Registra criacao", async () =>
    {
        File.WriteAllText(arquivo, "Arquivo usado no teste de criacao.");
        await Aguardar(() => LerLog().Contains("CRIADO | " + arquivo));
    });
    await Testar("Registra exclusao", async () =>
    {
        File.Delete(arquivo);
        await Aguardar(() => LerLog().Contains("EXCLUIDO | " + arquivo));
    });

    await Testar("Registra lote de 50 criacoes e 50 exclusoes", async () =>
    {
        var arquivos = Enumerable.Range(1, 50).Select(i => Path.Combine(entrada, $"lote-{i}.txt")).ToArray();
        foreach (var item in arquivos) File.WriteAllText(item, "Teste em lote.");
        await Aguardar(() =>
        {
            var texto = LerLog();
            return arquivos.All(item => texto.Contains("CRIADO | " + item + Environment.NewLine));
        });
        foreach (var item in arquivos) File.Delete(item);
        await Aguardar(() =>
        {
            var texto = LerLog();
            return arquivos.All(item => texto.Contains("EXCLUIDO | " + item + Environment.NewLine));
        });
    });
}
finally
{
    await worker.StopAsync(CancellationToken.None);
}
await Testar("Registra encerramento", () => Aguardar(() => LerLog().Contains("ENCERRADO | " + entrada)));
Console.WriteLine($"TOTAL=5 FALHAS={falhas}");
Console.WriteLine($"Log do teste: {caminhoLog}");
Environment.ExitCode = falhas == 0 ? 0 : 1;
