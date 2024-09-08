using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Reflection;
using System.Reflection.Emit;

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
            List<Type> dynamicEntityTypes = [];
            Dictionary<string, List<object>> tablesData = [];
            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var table in worksheet.Tables)
                {
                    Console.WriteLine($"Worksheet Name: {worksheet.Name}");
                    Console.WriteLine($"Table Name: {table.Name}");
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
                    var entityType = CreateDynamicType(tableName, tableHeaders, columnDataWithType.Select(c => c.dataType).ToList());
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

            var dbName = Path.GetFileNameWithoutExtension(filePath) + ".db";
            var dbPath = Path.Combine(directoryPath, dbName);
            if (File.Exists(dbPath))
            {
                File.Delete(dbPath);
            }

            using var context = new ExcelDbContext(dbPath, dynamicEntityTypes);
            context.Database.EnsureCreated();
            foreach (var tableData in tablesData)
            {
                context.AddRange(tableData.Value);
            }

            context.SaveChanges();
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

    static Type CreateDynamicType(string typeName, List<string> propertyNames, List<Type> propertyTypes)
    {
        var assemblyName = new AssemblyName("DynamicTypes");
        var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
        var moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
        var typeBuilder = moduleBuilder.DefineType(typeName, TypeAttributes.Public);

        for (int i = 0; i < propertyNames.Count; i++)
        {
            var fieldBuilder = typeBuilder.DefineField("_" + propertyNames[i], propertyTypes[i], FieldAttributes.Private);
            var propertyBuilder = typeBuilder.DefineProperty(propertyNames[i], PropertyAttributes.HasDefault, propertyTypes[i], null);

            var getterMethod = typeBuilder.DefineMethod("get_" + propertyNames[i], MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyTypes[i], Type.EmptyTypes);
            var getterIL = getterMethod.GetILGenerator();
            getterIL.Emit(OpCodes.Ldarg_0);
            getterIL.Emit(OpCodes.Ldfld, fieldBuilder);
            getterIL.Emit(OpCodes.Ret);

            var setterMethod = typeBuilder.DefineMethod("set_" + propertyNames[i], MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, null, [propertyTypes[i]]);
            var setterIL = setterMethod.GetILGenerator();
            setterIL.Emit(OpCodes.Ldarg_0);
            setterIL.Emit(OpCodes.Ldarg_1);
            setterIL.Emit(OpCodes.Stfld, fieldBuilder);
            setterIL.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getterMethod);
            propertyBuilder.SetSetMethod(setterMethod);

            // Add [Key] attribute to the PrimaryKey property
            if (propertyNames[i] == "PrimaryKey")
            {
                var keyAttributeConstructor = typeof(KeyAttribute).GetConstructor(Type.EmptyTypes);
                var keyAttributeBuilder = new CustomAttributeBuilder(keyAttributeConstructor!, []);
                propertyBuilder.SetCustomAttribute(keyAttributeBuilder);
            }
        }

        return typeBuilder.CreateType();
    }

    public static string ToUpperCamelCase(string input)
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

public class ExcelDbContext(string dbPath, List<Type> dynamicEntityTypes) : DbContext
{
    private readonly string dbPath = dbPath;
    private readonly List<Type> dynamicEntityTypes = dynamicEntityTypes;

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlite($"Data Source={dbPath}");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        foreach (var entityType in dynamicEntityTypes)
        {
            modelBuilder.Entity(entityType);
        }
    }
}
