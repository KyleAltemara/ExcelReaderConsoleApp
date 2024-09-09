using ExcelReaderConsoleApp.Interfaces;
using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

/// <summary>
/// Represents a class for reading Excel files.
/// </summary>
public class ExcelFileReader : IExcelFileReader
{
    /// <summary>
    /// Reads Excel files from the specified directory path.
    /// </summary>
    /// <param name="directoryPath">The directory path containing the Excel files.</param>
    /// <returns>A list of ExcelPackage objects representing the read Excel files. Only Excel files with tables are included.</returns>
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
