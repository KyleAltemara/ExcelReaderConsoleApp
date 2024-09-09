using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

public interface IExcelFileReader
{
    List<ExcelPackage> ReadExcelFiles(string directoryPath);
}
