using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for inserting rows in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class RowInsertionExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/RowInsertion";

        /// <summary>
        /// Initializes a new instance of the <see cref="RowInsertionExamples"/> class.
        /// Prepares the directory 'Documents/Excel/RowInsertion' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public RowInsertionExamples()
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
        /// Creates a new Excel workbook and demonstrates row insertion using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/RowInsertion' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "RowInsertion.xlsx").
        /// </param>
        public void CreateRowInsertion(string documentDirectory = docsDirectory, string filename = "RowInsertion.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a workbook first with some initial data
                using (var wb = new Workbook())
                {
                    var firstSheet = wb.Worksheets[0];

                    // Add some initial data to see the effect of row insertion
                    for (int i = 1; i <= 10; i++)
                    {
                        firstSheet.Cells[$"A{i}"].PutValue($"Row {i} - Column A");
                        firstSheet.Cells[$"B{i}"].PutValue($"Row {i} - Column B");
                        firstSheet.Cells[$"C{i}"].PutValue($"Row {i} - Column C");
                    }

                    wb.Save(filePath);
                }

                // Load the workbook from the specified file path
                using (var wb = new Workbook(filePath))
                {
                    // Access the first worksheet in the workbook
                    var firstSheet = wb.Worksheets[0];

                    // Define the starting row index and the number of rows to insert
                    uint startRowIndex = 5;
                    uint numberOfRows = 3;

                    // Insert the rows into the worksheet
                    firstSheet.InsertRows(startRowIndex, numberOfRows);

                    // Get the total row count after insertion
                    int rowsCount = firstSheet.GetRowCount();

                    // Output the updated row count to the console
                    Console.WriteLine("Rows Count=" + rowsCount);

                    // Save the workbook to reflect the changes made
                    wb.Save(filePath);

                    Console.WriteLine("Rows inserted and workbook saved successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating row insertion: {ex.Message}");
                throw;
            }
        }
    }
}