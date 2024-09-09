using ExcelReaderConsoleApp.Interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Globalization;

namespace ExcelReaderConsoleApp;

/// <summary>
/// Represents a data processor for processing Excel data.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref="DataProcessor"/> class.
/// </remarks>
/// <param name="typeBuilder">The dynamic type builder.</param>
/// <param name="logger">The logger of type <see cref="ILogger{DataProcessor}"/>.</param>
public class DataProcessor(DynamicTypeBuilder typeBuilder, ILogger<DataProcessor> logger) : IDataProcessor
{
    private readonly DynamicTypeBuilder _typeBuilder = typeBuilder ?? throw new ArgumentNullException(nameof(typeBuilder));

    private readonly ILogger<DataProcessor> _logger = logger ?? throw new ArgumentNullException(nameof(logger));

    /// <summary>
    /// Processes the data from the specified Excel package.
    /// </summary>
    /// <param name="package">The Excel package.</param>
    /// <returns>A tuple containing the dynamic entity types and the tables data.</returns>
    public (List<Type> dynamicEntityTypes, Dictionary<string, List<object>> tablesData) ProcessData(ExcelPackage package)
    {
        List<Type> dynamicEntityTypes = [];
        Dictionary<string, List<object>> tablesData = [];

        var workbook = package.Workbook;
        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var table in worksheet.Tables)
            {
                _logger.LogInformation("Worksheet Name: {worksheet}", worksheet.Name);
                _logger.LogInformation("Table Name: {table}", table.Name);
                var range = table.Range;
                List<string> tableHeaders = ["PrimaryKey"];
                for (int column = range.Start.Column; column <= range.End.Column; column++)
                {
                    tableHeaders.Add(ToUpperCamelCase(worksheet.Cells[range.Start.Row, column].Text));
                }

                List<List<string>> tableData = [Enumerable.Range(1, range.End.Row - range.Start.Row).Select(i => i.ToString()).ToList()];
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

                List<(Type dataType, List<object> data)> columnDataWithType = [];
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
                        else
                        {
                            data.Add(cellValue);
                        }
                    }
                }

                var tableName = $"{worksheet.Name}_{table.Name}";
                var entityType = _typeBuilder.CreateDynamicType(tableName, tableHeaders, columnDataWithType.Select(c => c.dataType).ToList());
                dynamicEntityTypes.Add(entityType);
                List<object> foo = [];
                tablesData.Add(tableName, foo);
                for (int row = 0; row < columnDataWithType[0].data.Count; row++)
                {
                    var entity = Activator.CreateInstance(entityType) ?? throw new InvalidOperationException($"Failed to create an instance of {entityType.Name}");
                    foo.Add(entity);
                    for (int col = 0; col < tableHeaders.Count; col++)
                    {
                        var property = entityType.GetProperty(tableHeaders[col]);
                        property?.SetValue(entity, columnDataWithType[col].data[row]);
                    }
                }
            }
        }

        return (dynamicEntityTypes, tablesData);
    }

    /// <summary>
    /// Itterates through the column data and determines the type of the column. If the column contains a mix of types, it will default to string.
    /// </summary>
    /// <param name="columnData"> The data in the column. </param>
    /// <returns> The type of the column. </returns>
    private static Type GetColumnType(List<string> columnData)
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
                if (columnType == typeof(double))
                {
                    continue;
                }

                columnType = typeof(int);
            }
            else if (double.TryParse(value, out _))
            {
                columnType = typeof(double);
            }
            else
            {
                columnType = typeof(string);
                break;
            }
        }

        return columnType is null ? typeof(string) : columnType;
    }

    /// <summary>
    /// Converts the input string to upper camel case.
    /// </summary>
    /// <param name="input"> The input string. </param>
    /// <returns> The input string in upper camel case. </returns>
    private static string ToUpperCamelCase(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }

        TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
        string[] words = input.Split([' ', '\t', '-', '_'], StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < words.Length; i++)
        {
            words[i] = textInfo.ToTitleCase(words[i].ToLower());
        }

        return string.Concat(words);
    }
}
