// See https://aka.ms/new-console-template for more information
using ExcelGenerator.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;

#region Configuration
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

#endregion Configuration

Console.WriteLine("Kindly create a list with values comma separated");
var text = Console.ReadLine();

try
{
    List<string> data = text.Split(',').ToList();
    var result = await services.GenerateAndReturnExcel(data);

    Console.WriteLine("The excel has been generated. Kindly find the content of the excel below:");
    Console.WriteLine(Convert.ToBase64String(result));
}
catch (Exception ex)
{
    Log.Logger.Error(ex, "Application caught an exception");
}


#region Methods
static void BuildConfig(IConfigurationBuilder configurationBuilder)
{
    configurationBuilder.SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
}

static List<PersonModel> GetListData()
{
    return new List<PersonModel> { new PersonModel() { Id = 1, Name = "Michael" },
                                   new PersonModel() { Id = 2, Name = "Fernandes" } };
}
#endregion Methods