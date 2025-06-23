using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for setting row heights and column widths in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class RowColumnExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/RowColumn";

        /// <summary>
        /// Initializes a new instance of the <see cref="RowColumnExamples"/> class.
        /// Prepares the directory 'Documents/Excel/RowColumn' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public RowColumnExamples()
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
        /// Creates a new Excel workbook with custom row heights and column widths using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/RowColumn' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "RowColumnSizing.xlsx").
        /// </param>
        public void CreateRowColumnSizing(string documentDirectory = docsDirectory, string filename = "RowColumnSizing.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Initialize a new workbook instance
                using (var wb = new Workbook())
                {
                    // Access the first worksheet in the workbook
                    var firstSheet = wb.Worksheets[0];

                    // Set the height of the first row to 40 points
                    firstSheet.SetRowHeight(1, 40);

                    // Set the width of column B to 75 points
                    firstSheet.SetColumnWidth("B", 75);

                    // Insert a value into cell A1
                    firstSheet.Cells["A1"].PutValue("Value in A1");

                    // Insert a styled text into cell B2
                    firstSheet.Cells["B2"].PutValue("Styled Text");

                    // Add some additional examples
                    firstSheet.SetRowHeight(2, 30);
                    firstSheet.SetColumnWidth("A", 50);
                    firstSheet.SetColumnWidth("C", 100);

                    firstSheet.Cells["A2"].PutValue("Row 2 - Height 30");
                    firstSheet.Cells["C1"].PutValue("Column C - Width 100");

                    // Save the workbook to the specified file path
                    wb.Save(filePath);

                    Console.WriteLine("Row and column sizing applied successfully and workbook saved.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating row/column sizing: {ex.Message}");
                throw;
            }
        }
    }
}