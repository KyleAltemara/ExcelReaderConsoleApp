namespace ExcelReaderConsoleApp.Interfaces;

public interface IDatabaseService
{
    void SaveData(string dbPath, List<Type> dynamicEntityTypes, Dictionary<string, List<object>> tablesData);
}
