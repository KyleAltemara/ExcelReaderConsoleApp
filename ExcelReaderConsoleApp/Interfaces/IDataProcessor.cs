using OfficeOpenXml;

namespace ExcelReaderConsoleApp.Interfaces;

public interface IDataProcessor
{
    (List<Type> dynamicEntityTypes, Dictionary<string, List<object>> tablesData) ProcessData(ExcelPackage package);
}
