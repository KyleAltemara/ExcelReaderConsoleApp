using ExcelReaderConsoleApp.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

internal class Program
{
    static void Main(string[] args)
    {
        var serviceProvider = new ServiceCollection()
            .AddLogging(configure => configure.AddConsole()) // Add logging services and configure console output
            .AddSingleton<DynamicTypeBuilder>() // Register the DynamicTypeBuilder service
            .AddSingleton<IExcelFileReader, ExcelFileReader>() // Register the IExcelFileReader service
            .AddSingleton<IDataProcessor, DataProcessor>() // Register the IDataProcessor service
            .AddSingleton<IDatabaseService, DatabaseService>() // Register the IDatabaseService service
            .BuildServiceProvider(); // Build the service provider

        // Set the logging level (optional)
        var logger = serviceProvider.GetService<ILogger<Program>>();
        logger?.LogInformation("Application started.");

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // License context required for EPPlus
        var directoryPath = "../../../../ExcelFiles";
        var excelFileReader = serviceProvider.GetService<IExcelFileReader>();
        var dataProcessor = serviceProvider.GetService<IDataProcessor>();
        var databaseService = serviceProvider.GetService<IDatabaseService>();
        if (excelFileReader is null || dataProcessor is null || databaseService is null)
        {
            logger?.LogError("Failed to resolve services.");
            return;
        }

        var packages = excelFileReader.ReadExcelFiles(directoryPath);
        foreach (var package in packages)
        {
            var (dynamicEntityTypes, tablesData) = dataProcessor.ProcessData(package);
            var dbPath = Path.Combine(package.File.DirectoryName!, $"{Path.GetFileNameWithoutExtension(package.File.Name)}.db");
            databaseService.SaveData(dbPath, dynamicEntityTypes, tablesData);
        }

        logger?.LogInformation("Application finished.");
    }
}
