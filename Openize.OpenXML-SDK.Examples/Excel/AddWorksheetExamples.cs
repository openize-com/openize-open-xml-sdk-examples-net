using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for adding worksheets to Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class AddWorksheetExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/AddWorksheet";

        /// <summary>
        /// Initializes a new instance of the <see cref="AddWorksheetExamples"/> class.
        /// Prepares the directory 'Documents/Excel/AddWorksheet' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public AddWorksheetExamples()
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
        /// Creates a new Excel workbook and demonstrates adding worksheets using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/AddWorksheet' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "AddWorksheet.xlsx").
        /// </param>
        public void CreateAddWorksheet(string documentDirectory = docsDirectory, string filename = "AddWorksheet.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a workbook first with some initial data
                using (var workbook = new Workbook())
                {
                    var firstSheet = workbook.Worksheets[0];

                    // Add some initial data to the first sheet
                    firstSheet.Cells["A1"].PutValue("This is the original sheet");
                    firstSheet.Cells["A2"].PutValue("Some initial data");

                    workbook.Save(filePath);
                }

                // Open the existing workbook.
                using (var workbook = new Workbook(filePath))
                {
                    // Add a new worksheet to the workbook.
                    var newSheet = workbook.AddSheet("NewWorksheetName"); // Replace 'NewWorksheetName' with the desired name

                    // You can also add some content to the new worksheet if needed.
                    var cellA1 = newSheet.Cells["A1"];
                    cellA1.PutValue("Hello from the new sheet!");

                    // Save the workbook with the added worksheet.
                    workbook.Save();

                    Console.WriteLine("New worksheet added and workbook saved successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating add worksheet: {ex.Message}");
                throw;
            }
        }
    }
}