using Microsoft.EntityFrameworkCore;

namespace ExcelReaderConsoleApp;

/// <summary>
/// Represents the Excel database context.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref="ExcelDbContext"/> class.
/// </remarks>
/// <param name="dbPath">The path to the SQLite database file.</param>
/// <param name="dynamicEntityTypes">The list of dynamic entity types. Each entity type represents a table in the database.</param>
public class ExcelDbContext(string dbPath, List<Type> dynamicEntityTypes) : DbContext
{
    private readonly string _dbPath = dbPath;
    private readonly List<Type> _dynamicEntityTypes = dynamicEntityTypes;

    /// <summary>
    /// Configures the database connection.
    /// </summary>
    /// <param name="optionsBuilder">The options builder.</param>
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlite($"Data Source={_dbPath}");
    }

    /// <summary>
    /// Configures the entity models.
    /// </summary>
    /// <param name="modelBuilder">The model builder.</param>
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // Register each dynamic entity type as a table in the database
        foreach (var entityType in _dynamicEntityTypes)
        {
            modelBuilder.Entity(entityType);
        }
    }
}
