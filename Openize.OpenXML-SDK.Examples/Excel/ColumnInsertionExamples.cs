using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for inserting columns in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class ColumnInsertionExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/ColumnInsertion";

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnInsertionExamples"/> class.
        /// Prepares the directory 'Documents/Excel/ColumnInsertion' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public ColumnInsertionExamples()
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
        /// Creates a new Excel workbook and demonstrates column insertion using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/ColumnInsertion' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "ColumnInsertion.xlsx").
        /// </param>
        public void CreateColumnInsertion(string documentDirectory = docsDirectory, string filename = "ColumnInsertion.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a workbook first with some initial data
                using (var wb = new Workbook())
                {
                    var firstSheet = wb.Worksheets[0];

                    // Add some initial data to see the effect of column insertion
                    firstSheet.Cells["A1"].PutValue("Column A");
                    firstSheet.Cells["B1"].PutValue("Column B");
                    firstSheet.Cells["C1"].PutValue("Column C");
                    firstSheet.Cells["D1"].PutValue("Column D");

                    firstSheet.Cells["A2"].PutValue("Data A2");
                    firstSheet.Cells["B2"].PutValue("Data B2");
                    firstSheet.Cells["C2"].PutValue("Data C2");
                    firstSheet.Cells["D2"].PutValue("Data D2");

                    wb.Save(filePath);
                }

                // Load and insert columns following the example pattern
                using (var wb = new Workbook(filePath))
                {
                    string startColumn = "B";
                    int numberOfColumns = 3;

                    var firstSheet = wb.Worksheets[0];

                    int columnsCount = firstSheet.GetColumnCount();
                    firstSheet.InsertColumns(startColumn, numberOfColumns);

                    Console.WriteLine("Columns Count=" + columnsCount);

                    wb.Save(filePath);

                    Console.WriteLine("Columns inserted successfully and workbook saved.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating column insertion: {ex.Message}");
                throw;
            }
        }
    }
}