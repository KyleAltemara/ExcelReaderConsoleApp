using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

public class ExcelFileReader : IExcelFileReader
{
    public List<ExcelPackage> ReadExcelFiles(string directoryPath)
    {
        List<ExcelPackage> packages = [];
        string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");
        foreach (var filePath in excelFiles)
        {
            if (Path.GetFileName(filePath).StartsWith("~$"))
            {
                continue;
            }

            FileInfo fileInfo = new(filePath);
            ExcelPackage package = new(fileInfo);
            if (package is null ||
                package.Workbook is null ||
                package.Workbook.Worksheets is null ||
                package.Workbook.Worksheets.Count == 0 ||
                package.Workbook.Worksheets.All(w => w.Tables.Count == 0))
            {
                continue;
            }

            packages.Add(package);
        }

        return packages;
    }
}
