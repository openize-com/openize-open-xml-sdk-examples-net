using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for cell styling in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class CellStylingExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/CellStyling";

        /// <summary>
        /// Initializes a new instance of the <see cref="CellStylingExamples"/> class.
        /// Prepares the directory 'Documents/Excel/CellStyling' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public CellStylingExamples()
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
        /// Creates a new Excel workbook and demonstrates cell styling using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/CellStyling' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "CellStyling.xlsx").
        /// </param>
        public void CreateCellStyling(string documentDirectory = docsDirectory, string filename = "CellStyling.xlsx")
        {
            try
            {
                // Define the path to save the spreadsheet
                string filePath = $"{documentDirectory}/{filename}";

                // Creating a new workbook instance
                using (var wb = new Workbook())
                {
                    // Create a custom style with Arial font, size 11, and red color
                    uint styleIndex = wb.CreateStyle("Arial", 11, "FF0000");

                    // Create another custom style with Times New Roman font, size 12, and black color
                    uint styleIndex2 = wb.CreateStyle("Times New Roman", 12, "000000");

                    // Access the first worksheet from the workbook
                    var firstSheet = wb.Worksheets[0];

                    // Assign a value to the cell A1 and apply the first custom style
                    var cellA1 = firstSheet.Cells["A1"];
                    cellA1.PutValue("Styled Text A1");
                    cellA1.ApplyStyle(styleIndex);

                    // Assign a value to the cell B2 and apply the second custom style
                    var cellB2 = firstSheet.Cells["B2"];
                    cellB2.PutValue("Styled Text B2");
                    cellB2.ApplyStyle(styleIndex2);

                    // Save the workbook to the specified file path
                    wb.Save(filePath);

                    Console.WriteLine("Cell styling applied and workbook saved successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating cell styling: {ex.Message}");
                throw;
            }
        }
    }
}