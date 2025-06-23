using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for working with cell ranges in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class RangeExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/Range";

        /// <summary>
        /// Initializes a new instance of the <see cref="RangeExamples"/> class.
        /// Prepares the directory 'Documents/Excel/Range' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public RangeExamples()
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
        /// Creates a new Excel workbook and demonstrates range operations using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/Range' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "RangeExample.xlsx").
        /// </param>
        public void CreateRangeExample(string documentDirectory = docsDirectory, string filename = "RangeExample.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a workbook first with some initial data
                using (var wb = new Workbook())
                {
                    var firstSheet = wb.Worksheets[0];

                    // Add some initial data
                    firstSheet.Cells["A1"].PutValue("Initial Data");
                    firstSheet.Cells["B1"].PutValue("Will be replaced");

                    wb.Save(filePath);
                }

                // Load the workbook from the specified file path
                using (var wb = new Workbook(filePath))
                {
                    // Access the first worksheet in the workbook
                    var firstSheet = wb.Worksheets[0];

                    // Select a range within the worksheet
                    var range = firstSheet.GetRange("A1", "B10");
                    Console.WriteLine($"Column count: {range.ColumnCount}");
                    Console.WriteLine($"Row count: {range.RowCount}");

                    // Set a similar value to all cells in the selected range
                    range.SetValue("Hello");

                    // Save the changes back to the workbook
                    wb.Save(filePath);

                    Console.WriteLine("Value set to range and workbook saved successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating range example: {ex.Message}");
                throw;
            }
        }
    }
}