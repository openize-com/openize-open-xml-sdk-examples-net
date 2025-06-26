using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for freezing panes in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class FreezePanesExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/FreezePanes";

        /// <summary>
        /// Initializes a new instance of the <see cref="FreezePanesExamples"/> class.
        /// Prepares the directory 'Documents/Excel/FreezePanes' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public FreezePanesExamples()
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
        /// Creates a new Excel workbook with frozen panes using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/FreezePanes' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "FrozenPanes.xlsx").
        /// </param>
        public void CreateFreezePanes(string documentDirectory = docsDirectory, string filename = "FrozenPanes.xlsx")
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

                    // Freeze the first column
                    worksheet.FreezePane(0, 1);

                    // Save the workbook
                    workbook.Save(filePath);
                    Console.WriteLine("Excel file created with frozen first column!");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating freeze panes: {ex.Message}");
                throw;
            }
        }
    }
}