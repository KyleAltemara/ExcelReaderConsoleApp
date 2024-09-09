using Microsoft.Extensions.Logging;

namespace ExcelReaderConsoleApp;

public class DatabaseService : IDatabaseService
{
    private readonly ILogger<DatabaseService> _logger;

    public DatabaseService(ILogger<DatabaseService> logger)
    {
        _logger = logger;
    }

    public void SaveData(string dbPath, List<Type> dynamicEntityTypes, Dictionary<string, List<object>> tablesData)
    {
        if (File.Exists(dbPath))
        {
            File.Delete(dbPath);
        }

        var typeBuilder = new DynamicTypeBuilder();
        var contextType = typeBuilder.CreateInheritedType(dbPath, typeof(ExcelDbContext), [typeof(string), typeof(List<Type>)]);
        using (var context = Activator.CreateInstance(contextType, dbPath, dynamicEntityTypes) as ExcelDbContext)
        {
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
                context?.Dispose();
                _logger.LogInformation("ExcelDbContext disposed.");
            }
        }
    }
}
