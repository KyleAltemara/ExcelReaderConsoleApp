using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

public interface IDataProcessor
{
    (List<Type> dynamicEntityTypes, Dictionary<string, List<object>> tablesData) ProcessData(ExcelPackage package);
}
