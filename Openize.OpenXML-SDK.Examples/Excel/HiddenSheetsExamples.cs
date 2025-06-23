using System;
using System.Collections.Generic;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for working with hidden sheets in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class HiddenSheetsExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/HiddenSheets";

        /// <summary>
        /// Initializes a new instance of the <see cref="HiddenSheetsExamples"/> class.
        /// Prepares the directory 'Documents/Excel/HiddenSheets' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public HiddenSheetsExamples()
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
        /// Creates a workbook with multiple sheets and hides some of them using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/HiddenSheets' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "HiddenSheets.xlsx").
        /// </param>
        public void CreateHiddenSheets(string documentDirectory = docsDirectory, string filename = "HiddenSheets.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a workbook with multiple sheets first
                using (var wb = new Workbook())
                {
                    // Add some additional sheets
                    var sheet2 = wb.AddSheet("TestSheet");
                    var sheet3 = wb.AddSheet("DataSheet");
                    var sheet4 = wb.AddSheet("ReportSheet");

                    // Add some content to sheets
                    wb.Worksheets[0].Cells["A1"].PutValue("Main Sheet");
                    sheet2.Cells["A1"].PutValue("Test Sheet Content");
                    sheet3.Cells["A1"].PutValue("Data Sheet Content");
                    sheet4.Cells["A1"].PutValue("Report Sheet Content");

                    wb.Save(filePath);
                }

                string sheetName = "TestSheet";

                // Load the workbook from the specified file path
                using (var wb = new Workbook(filePath))
                {
                    wb.SetSheetVisibility(sheetName, SheetVisibility.Hidden);
                    wb.Save(filePath);

                    Console.WriteLine($"Sheet '{sheetName}' has been hidden successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating hidden sheets: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Retrieves and displays information about hidden sheets in an Excel workbook using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook is located (default is the 'Documents/Excel/HiddenSheets' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file to read (default is "HiddenSheets.xlsx").
        /// </param>
        public void GetHiddenSheets(string documentDirectory = docsDirectory, string filename = "HiddenSheets.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Check if the file exists, create it if not
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine("Creating the workbook with hidden sheets first...");
                    CreateHiddenSheets(documentDirectory, filename);
                }

                // Specify the path to the Excel file with hidden sheets
                using (var wb = new Workbook(filePath))
                {
                    List<Tuple<string, string>> hiddenSheets = wb.GetHiddenSheets();

                    // Display information about hidden sheets
                    Console.WriteLine($"Found {hiddenSheets.Count} hidden sheets in {filename}:");

                    foreach (var sheet in hiddenSheets)
                    {
                        Console.WriteLine($"Hidden Sheet ID: {sheet.Item1}, Name: {sheet.Item2}");
                    }

                    Console.WriteLine("\nProcessing complete!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting hidden sheets: {ex.Message}");
                throw;
            }
        }
    }
}