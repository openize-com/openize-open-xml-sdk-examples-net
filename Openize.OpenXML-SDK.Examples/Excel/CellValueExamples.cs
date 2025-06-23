using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for basic cell value operations in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class CellValueExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/CellValue";

        /// <summary>
        /// Initializes a new instance of the <see cref="CellValueExamples"/> class.
        /// Prepares the directory 'Documents/Excel/CellValue' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public CellValueExamples()
        {
            if (!Directory.Exists(docsDirectory))
            {
                // If it doesn't exist, create the directory
                Directory.CreateDirectory(docsDirectory);
                Console.WriteLine($"Directory '{Path.GetFullPath(docsDirectory)}' created successfully.");
            }
            else
            {
                var files = Directory.GetFiles(Path.GetFullPath(docsDirectory));
                foreach (var file in files)
                {
                    File.Delete(file);
                    Console.WriteLine($"File deleted: {file}");
                }
                Console.WriteLine($"Directory '{Path.GetFullPath(docsDirectory)}' cleaned up.");
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with basic cell values using Openize.OpenXML-SDK.
        /// Demonstrates setting simple values in different cells.
        /// Based on: https://gist.github.com/openize-com-gists/7e88d8b52c383c6fe29a0ed89afb71ca
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/CellValue' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "CellValues.xlsx").
        /// </param>
        public void CreateCellValues(string documentDirectory = docsDirectory, string filename = "CellValues.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    // Get the first worksheet
                    var worksheet = workbook.Worksheets[0];

                    // Add some sample data
                    worksheet.Cells["A1"].PutValue("Employee ID");
                    worksheet.Cells["B1"].PutValue("First Name");
                    worksheet.Cells["C1"].PutValue("Last Name");
                    worksheet.Cells["D1"].PutValue("Department");
                    worksheet.Cells["E1"].PutValue("Salary");

                    // Add sample employee data
                    for (int i = 2; i <= 10; i++)
                    {
                        worksheet.Cells[$"A{i}"].PutValue($"EMP{i:000}");
                        worksheet.Cells[$"B{i}"].PutValue($"FirstName{i}");
                        worksheet.Cells[$"C{i}"].PutValue($"LastName{i}");
                        worksheet.Cells[$"D{i}"].PutValue($"Department{i % 3 + 1}");
                        worksheet.Cells[$"E{i}"].PutValue(50000 + (i * 1000));
                    }

                    // Save the workbook
                    workbook.Save(filePath);
                    Console.WriteLine("Excel file created successfully");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Reads cell values from an existing Excel workbook using Openize.OpenXML-SDK.
        /// Demonstrates how to access and retrieve values from different cells.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook is located (default is the 'Documents/Excel/CellValue' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file to read (default is "CellValues.xlsx").
        /// </param>
        public void ReadCellValues(string documentDirectory = docsDirectory, string filename = "CellValues.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Check if the file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine("Creating the workbook first...");
                    CreateCellValues(documentDirectory, filename);
                }

                // Load the workbook from the specified file path
                using (var workbook = new Workbook(filePath))
                {
                    // Get the first worksheet
                    var worksheet = workbook.Worksheets[0];

                    Console.WriteLine($"Reading values from worksheet: {worksheet.Name}");
                    Console.WriteLine("=====================================");

                    // Read header row
                    Console.WriteLine("Headers:");
                    for (char col = 'A'; col <= 'E'; col++)
                    {
                        var headerValue = worksheet.Cells[$"{col}1"].GetValue();
                        Console.WriteLine($"  {col}1: {headerValue}");
                    }

                    Console.WriteLine("\nEmployee Data:");
                    Console.WriteLine("----------------------------------");

                    // Read employee data (rows 2-10)
                    for (int row = 2; row <= 10; row++)
                    {
                        var empId = worksheet.Cells[$"A{row}"].GetValue();
                        var firstName = worksheet.Cells[$"B{row}"].GetValue();
                        var lastName = worksheet.Cells[$"C{row}"].GetValue();
                        var department = worksheet.Cells[$"D{row}"].GetValue();
                        var salary = worksheet.Cells[$"E{row}"].GetValue();

                        Console.WriteLine($"Row {row}: {empId} | {firstName} {lastName} | {department} | ${salary}");
                    }

                    // Show some specific cell access examples
                    Console.WriteLine("\nSpecific Cell Access:");
                    Console.WriteLine("--------------------");
                    Console.WriteLine($"Employee ID in A2: {worksheet.Cells["A2"].GetValue()}");
                    Console.WriteLine($"First salary in E2: {worksheet.Cells["E2"].GetValue()}");
                    Console.WriteLine($"Last employee in A10: {worksheet.Cells["A10"].GetValue()}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading cell values: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Updates cell values in an existing Excel workbook using Openize.OpenXML-SDK.
        /// Demonstrates modifying existing cell values and adding new data.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook is located and where the modified workbook will be saved.
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file to modify (default is "CellValues.xlsx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Excel workbook (default is "UpdatedCellValues.xlsx").
        /// </param>
        public void UpdateCellValues(string documentDirectory = docsDirectory,
            string filename = "CellValues.xlsx", string filenameModified = "UpdatedCellValues.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";
                string modifiedFilePath = $"{documentDirectory}/{filenameModified}";

                // Check if the file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine("Creating the workbook first...");
                    CreateCellValues(documentDirectory, filename);
                }

                // Load the workbook from the specified file path
                using (var workbook = new Workbook(filePath))
                {
                    // Get the first worksheet
                    var worksheet = workbook.Worksheets[0];

                    Console.WriteLine($"Updating values in worksheet: {worksheet.Name}");

                    // Update some employee salaries (give raises!)
                    Console.WriteLine("Giving salary raises to employees...");
                    for (int row = 2; row <= 10; row++)
                    {
                        var currentSalary = Convert.ToInt32(worksheet.Cells[$"E{row}"].GetValue());
                        var newSalary = currentSalary + 5000; // $5000 raise
                        worksheet.Cells[$"E{row}"].PutValue(newSalary);

                        var empId = worksheet.Cells[$"A{row}"].GetValue();
                        Console.WriteLine($"  {empId}: ${currentSalary} → ${newSalary}");
                    }

                    // Add a new employee (row 11)
                    Console.WriteLine("\nAdding new employee...");
                    worksheet.Cells["A11"].PutValue("EMP011");
                    worksheet.Cells["B11"].PutValue("NewFirstName");
                    worksheet.Cells["C11"].PutValue("NewLastName");
                    worksheet.Cells["D11"].PutValue("Department1");
                    worksheet.Cells["E11"].PutValue(75000);
                    Console.WriteLine("  Added: EMP011 | NewFirstName NewLastName | Department1 | $75000");

                    // Update a department name
                    Console.WriteLine("\nUpdating department names...");
                    for (int row = 2; row <= 11; row++)
                    {
                        var currentDept = worksheet.Cells[$"D{row}"].GetValue().ToString();
                        if (currentDept == "Department1")
                        {
                            worksheet.Cells[$"D{row}"].PutValue("HR Department");
                            var empId = worksheet.Cells[$"A{row}"].GetValue();
                            Console.WriteLine($"  {empId}: Department1 → HR Department");
                        }
                    }

                    // Add a timestamp
                    worksheet.Cells["G1"].PutValue("Last Updated");
                    worksheet.Cells["G2"].PutValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    // Save the modified workbook
                    workbook.Save(modifiedFilePath);
                    Console.WriteLine($"\nModified workbook saved as {filenameModified}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating cell values: {ex.Message}");
                throw;
            }
        }


    }
}