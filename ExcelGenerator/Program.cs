// See https://aka.ms/new-console-template for more information
using ExcelGenerator.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;

Console.WriteLine("Hello, World!");

var builder = new ConfigurationBuilder();
BuildConfig(builder);

Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Build())
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .CreateLogger();

Log.Logger.Information("Application is starting");

var host = Host.CreateDefaultBuilder()
    .ConfigureServices((context, services) =>
    {
        services.AddTransient<IExcelService, ExcelService>();
    })
    .UseSerilog()
    .Build();

var services = ActivatorUtilities.CreateInstance<ExcelService>(host.Services);
await services.GenerateExcel();

static void BuildConfig(IConfigurationBuilder configurationBuilder)
{
    configurationBuilder.SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
}