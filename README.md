# ExcelReaderConsoleApp

This is a console application that reads Excel files, processes the data, and saves it to a SQLite database. The application uses dynamic types to handle various Excel table structures and displays the data using Spectre.Console.

<https://www.thecsharpacademy.com/project/20/excel-reader>

## Features

- Reads Excel files from a specified directory.
- Processes Excel data and dynamically creates entity types.
- Saves the processed data to a SQLite database.
- Displays the data from the SQLite database in a user-friendly console interface.
- Uses Entity Framework Core to interact with the database and create the necessary schema.
- Uses Spectre.Console to create a user-friendly console interface.

## Getting Started

To run the application, follow these steps:

1. Clone the repository to your local machine.
2. Open the solution in Visual Studio.
3. Build the solution to restore NuGet packages and compile the code.
4. Run the `ExcelReaderConsoleApp` project to start the console application.

## Dependencies

- Microsoft.EntityFrameworkCore: The application uses this package to manage the database context and entity relationships.
- Microsoft.EntityFrameworkCore.Sqlite: The application uses this package to interact with a SQLite database.
- Spectre.Console: The application uses this package to create a user-friendly console interface.
- EPPlus: The application uses this package to read and process Excel files.
- Microsoft.Data.Sqlite: The application uses this package to interact with a SQLite database.

## Usage

1. The application will read Excel files from the specified directory.
2. It will process the data and dynamically create entity types based on the Excel table structures.
3. The processed data will be saved to a SQLite database.
4. The application will display the data from the database in a user-friendly console interface.

## License

This project is licensed under the MIT License.

## Resources Used

- [The C# Academy](https://www.thecsharpacademy.com/)
- [Sample Data - Contextures](https://www.contextures.com/xlsampledata01.html)
- [Sample Data - Food Sales](https://www.contextures.com/excelsampledatafoodsales.html)
- [Sample Data - Athletes](https://www.contextures.com/excelsampledataathletes.html)
- [Sample Data - Food Info](https://www.contextures.com/excelsampledatafoodinfo.html)
- [EPPlus Software](https://epplussoftware.com/)
- GitHub Copilot to generate code snippets.
