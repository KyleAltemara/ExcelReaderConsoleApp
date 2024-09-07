using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

internal class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Load the Excel package
        var directoryPath = "../../../../ExcelFiles";
        string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");
        foreach (var filePath in excelFiles)
        {
            FileInfo fileInfo = new(filePath);

            using ExcelPackage package = new(fileInfo);
            // Get the workbook
            var workbook = package.Workbook;

            // Iterate through all named ranges
            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var table in worksheet.Tables)
                {
                    Console.WriteLine($"Worksheet Name: {worksheet.Name}");
                    Console.WriteLine($"Table Name: {table.Name}");
                    var range = table.Range;
                    List<string> tableHeaders = [];
                    Console.Write("Table Headers: ");
                    for (int column = range.Start.Column; column <= range.End.Column; column++)
                    {
                        tableHeaders.Add(worksheet.Cells[range.Start.Row, column].Text);
                        Console.Write($"{worksheet.Cells[range.Start.Row, column].Text} ");
                    }

                    Console.WriteLine();
                    List<List<object>> tableData = [];

                    // Iterate through each row in the table
                    for (int row = range.Start.Row + 1; row <= range.End.Row; row++)
                    {
                        List<object> rowData = [];

                        // Iterate through each column in the row
                        for (int column = range.Start.Column; column <= range.End.Column; column++)
                        {
                            var cellValue = worksheet.Cells[row, column].Value;
                            rowData.Add(cellValue);
                        }

                        tableData.Add(rowData);
                    }

                    // Use the tableData list for further processing
                }
            }
        }
    }
}
