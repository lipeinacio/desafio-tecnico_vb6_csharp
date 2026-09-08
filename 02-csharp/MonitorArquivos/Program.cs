using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = Host.CreateApplicationBuilder(args);
builder.Services.AddWindowsService(options => options.ServiceName = "PsMonitorArquivos");
builder.Services.AddHostedService<MonitorWorker>();
await builder.Build().RunAsync();
