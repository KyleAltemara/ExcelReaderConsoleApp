using ExcelReaderConsoleApp.Interfaces;
using Microsoft.Extensions.Logging;

namespace ExcelReaderConsoleApp;

/// <summary>
/// Represents a service for saving data to a database.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref="DatabaseService"/> class.
/// </remarks>
/// <param name="logger">The logger of type <see cref="ILogger{DatabaseService}"/>.</param>
public class DatabaseService(ILogger<DatabaseService> logger) : IDatabaseService
{
    private readonly ILogger<DatabaseService> _logger = logger;

    /// <summary>
    /// Saves the data to the database.
    /// </summary>
    /// <param name="dbPath">The path of the database file.</param>
    /// <param name="dynamicEntityTypes">The dynamic entity types.</param>
    /// <param name="tablesData">The tables data.</param>
    public void SaveData(string dbPath, List<Type> dynamicEntityTypes, Dictionary<string, List<object>> tablesData)
    {
        if (File.Exists(dbPath))
        {
            File.Delete(dbPath);
        }

        var typeBuilder = new DynamicTypeBuilder();
        var contextType = typeBuilder.CreateInheritedType(dbPath, typeof(ExcelDbContext), [typeof(string), typeof(List<Type>)]);
        using var context = Activator.CreateInstance(contextType, dbPath, dynamicEntityTypes) as ExcelDbContext;
        try
        {
            context!.Database.EnsureCreated();
            foreach (var tableData in tablesData)
            {
                context.AddRange(tableData.Value);
            }

            context.SaveChanges();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "An error occurred while saving data to the database.");
            throw;
        }
        finally
        {
            _logger.LogInformation("Data saved to database: {dbPath}", dbPath);
            context?.Dispose();
        }
    }
}
