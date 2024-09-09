using OfficeOpenXml;

namespace ExcelReaderConsoleApp.Interfaces;

public interface IExcelFileReader
{
    List<ExcelPackage> ReadExcelFiles(string directoryPath);
}
