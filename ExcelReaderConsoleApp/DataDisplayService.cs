using ExcelReaderConsoleApp.Interfaces;
using Microsoft.Data.Sqlite;
using Spectre.Console;

namespace ExcelReaderConsoleApp;

/// <summary>
/// Represents a service for displaying data from an SQLite database.
/// </summary>
public class DataDisplayService : IDataDisplayService
{
    /// <summary>
    /// Displays the data from the SQLite database.
    /// </summary>
    /// <param name="dbPath">The path to the SQLite database file.</param>
    public void DisplayData(string dbPath)
    {
        using var connection = new SqliteConnection($"Data Source={dbPath}");
        connection.Open();

        // Get the table names. sqlite_master is a system table that defines the schema of the database.
        var command = new SqliteCommand("SELECT name FROM sqlite_master WHERE type='table';", connection);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var tableName = reader.GetString(0);
            if (tableName == "sqlite_sequence")
            {
                continue; // Skip the sqlite_sequence table
            }

            AnsiConsole.MarkupLine($"[bold yellow]Table: {tableName}[/]");

            // Get the number of columns. PRAGMA table_info returns one row for each column in the table.
            var tableCommand = new SqliteCommand($"PRAGMA table_info({tableName});", connection);
            using var tableReader = tableCommand.ExecuteReader();
            int columnCount = 0;
            while (tableReader.Read())
            {
                columnCount++;
            }

            AnsiConsole.MarkupLine($"[bold green]Number of columns: {columnCount}[/]");

            // Get the number of rows
            tableCommand = new SqliteCommand($"SELECT COUNT(*) FROM {tableName};", connection);
            int rowCount = Convert.ToInt32(tableCommand.ExecuteScalar());
            AnsiConsole.MarkupLine($"[bold green]Number of rows: {rowCount}[/]");

            // Display the table data
            tableCommand = new SqliteCommand($"SELECT * FROM {tableName};", connection);
            using var dataReader = tableCommand.ExecuteReader();

            var table = new Table();
            for (int i = 0; i < dataReader.FieldCount; i++)
            {
                table.AddColumn(new TableColumn(dataReader.GetName(i)));
            }

            while (dataReader.Read())
            {
                var row = new List<string>();
                for (int i = 0; i < dataReader.FieldCount; i++)
                {
                    row.Add(dataReader.GetValue(i)?.ToString() ?? "");
                }

                table.AddRow(row.ToArray());
            }

            AnsiConsole.Write(table);
        }
    }
}
