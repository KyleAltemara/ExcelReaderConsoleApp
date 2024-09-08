using Microsoft.EntityFrameworkCore;

namespace ExcelReaderConsoleApp;

public class ExcelDbContext(string dbPath, List<Type> dynamicEntityTypes) : DbContext
{
    private readonly string _dbPath = dbPath;
    private readonly List<Type> _dynamicEntityTypes = dynamicEntityTypes;

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlite($"Data Source={_dbPath}");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        foreach (var entityType in _dynamicEntityTypes)
        {
            modelBuilder.Entity(entityType);
        }
    }
}
