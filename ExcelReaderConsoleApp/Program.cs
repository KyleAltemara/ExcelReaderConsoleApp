using OfficeOpenXml;

namespace ExcelReaderConsoleApp;

internal class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var directoryPath = "../../../../ExcelFiles";
        string[] excelFiles = Directory.GetFiles(directoryPath, "*.xlsx");
        foreach (var filePath in excelFiles)
        {
            if (Path.GetFileName(filePath).StartsWith("~$"))
            {
                continue;
            }

            FileInfo fileInfo = new(filePath);
            using ExcelPackage package = new(fileInfo);
            var workbook = package.Workbook;

            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var table in worksheet.Tables)
                {
                    Console.WriteLine($"Worksheet Name: {worksheet.Name}");
                    Console.WriteLine($"Table Name: {table.Name}");
                    var range = table.Range;
                    List<string> tableHeaders = [];
                    for (int column = range.Start.Column; column <= range.End.Column; column++)
                    {
                        tableHeaders.Add(worksheet.Cells[range.Start.Row, column].Text);
                    }

                    List<List<string>> tableData = [];
                    for (int col = range.Start.Column; col <= range.End.Column; col++)
                    {
                        List<string> columnData = [];

                        for (int row = range.Start.Row + 1; row <= range.End.Row; row++)
                        {
                            var cellValue = worksheet.Cells[row, col].Text;
                            columnData.Add(cellValue);
                        }

                        tableData.Add(columnData);
                    }

                    List<(Type, List<object>)> columnDataWithType = [];
                    for (int col = 0; col < tableHeaders.Count; col++)
                    {
                        var columnType = GetColumnType(tableData[col]);
                        List<object> data = [];
                        columnDataWithType.Add((columnType, data));
                        for (int row = 0; row < tableData[col].Count; row++)
                        {
                            var cellValue = tableData[col][row];
                            if (columnType == typeof(int))
                            {
                                if (string.IsNullOrWhiteSpace(cellValue))
                                {
                                    data.Add(0);
                                }
                                else
                                {
                                    data.Add(int.Parse(cellValue));
                                }
                            }
                            else if (columnType == typeof(double))
                            {
                                if (string.IsNullOrWhiteSpace(cellValue))
                                {
                                    data.Add(0.0);
                                }
                                else
                                {
                                    data.Add(double.Parse(cellValue));
                                }
                            }
                            else if (columnType == typeof(DateTime))
                            {
                                if (string.IsNullOrWhiteSpace(cellValue))
                                {
                                    data.Add(DateTime.MinValue);
                                }
                                else
                                {
                                    data.Add(DateTime.Parse(cellValue));
                                }
                            }
                            else
                            {
                                data.Add(cellValue);
                            }
                        }
                    }

                    var tableName = $"{worksheet.Name}_{table.Name}";

                }
            }
        }
    }

    static Type GetColumnType(List<string> columnData)
    {
        Type? columnType = null;

        foreach (var value in columnData)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            if (int.TryParse(value, out _))
            {
                if (columnType == typeof(double) || columnType != typeof(DateTime))
                {
                    continue;
                }

                columnType = typeof(int);
            }
            else if (double.TryParse(value, out _))
            {
                if (columnType == typeof(DateTime))
                {
                    continue;
                }

                columnType = typeof(double);
            }
            else if (DateTime.TryParse(value, out _))
            {
                columnType = typeof(DateTime);
            }
            else
            {
                columnType = typeof(string);
                break;
            }

        }

        return columnType is null ? typeof(string) : columnType;
    }
}