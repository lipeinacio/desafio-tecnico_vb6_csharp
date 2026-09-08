using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

public sealed class MonitorWorker : IHostedService, IDisposable
{
    private readonly ILogger<MonitorWorker> logger;
    private readonly string pastaEntrada;
    private readonly string pastaLogs;
    private readonly object trava = new object();
    private FileSystemWatcher? monitor;

    public MonitorWorker(ILogger<MonitorWorker> logger, IConfiguration configuracao)
    {
        this.logger = logger;
        pastaEntrada = Path.GetFullPath(
            configuracao["Monitoramento:PastaEntrada"] ?? "monitoramento/entrada",
            AppContext.BaseDirectory);
        pastaLogs = Path.GetFullPath(
            configuracao["Monitoramento:PastaLogs"] ?? "monitoramento/logs",
            AppContext.BaseDirectory);
    }

    // O Windows chama este metodo quando o servico inicia.
    public Task StartAsync(CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(pastaEntrada);
        Directory.CreateDirectory(pastaLogs);

        monitor = new FileSystemWatcher(pastaEntrada);
        monitor.NotifyFilter = NotifyFilters.FileName;
        monitor.IncludeSubdirectories = false;
        monitor.Created += ArquivoCriado;
        monitor.Deleted += ArquivoExcluido;
        monitor.Error += ErroNoMonitor;
        monitor.EnableRaisingEvents = true;

        Registrar("INICIADO", pastaEntrada);
        return Task.CompletedTask;
    }

    private void ArquivoCriado(object sender, FileSystemEventArgs e)
    {
        Registrar("CRIADO", e.FullPath);
    }

    private void ArquivoExcluido(object sender, FileSystemEventArgs e)
    {
        Registrar("EXCLUIDO", e.FullPath);
    }

    private void ErroNoMonitor(object sender, ErrorEventArgs e)
    {
        logger.LogError(e.GetException(), "Falha no monitoramento.");
        Registrar("ERRO_MONITOR", e.GetException().Message);
    }

    private void Registrar(string evento, string arquivo)
    {
        // Evita duas gravacoes simultaneas no mesmo log.
        lock (trava)
        {
            try
            {
                string linha = $"{DateTimeOffset.Now:O} | {evento} | {arquivo}";
                File.AppendAllText(Path.Combine(pastaLogs, "eventos.log"),
                    linha + Environment.NewLine, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Nao foi possivel gravar o arquivo .log.");
            }
        }
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        if (monitor != null)
        {
            Dispose();
            Registrar("ENCERRADO", pastaEntrada);
        }
        return Task.CompletedTask;
    }

    public void Dispose()
    {
        if (monitor == null) return;
        monitor.EnableRaisingEvents = false;
        monitor.Dispose();
        monitor = null;
    }
}
