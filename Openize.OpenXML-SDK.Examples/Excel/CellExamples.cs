using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for working with Excel cells
    /// using the Openize.OpenXML-SDK library.
    /// </summary>
    public class CellExamples
    {
        private const string docsDirectory = "../../../Documents/Cell";

        /// <summary>
        /// Initializes a new instance of the <see cref="CellExamples"/> class.
        /// Prepares the directory 'Documents/Cell' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public CellExamples()
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
        /// Creates a new Excel workbook with cells containing different data types using Openize.OpenXML-SDK.
        /// Demonstrates how to set string, numeric, date, and boolean values in cells.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Cell' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "CellDataTypes.xlsx").
        /// </param>
        public void CreateCellsWithDifferentDataTypes(string documentDirectory = docsDirectory, string filename = "CellDataTypes.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];

                    // Add headers
                    worksheet.Cells["A1"].PutValue("Data Type");
                    worksheet.Cells["B1"].PutValue("Value");
                    worksheet.Cells["C1"].PutValue("Description");

                    // String values
                    worksheet.Cells["A2"].PutValue("String");
                    worksheet.Cells["B2"].PutValue("Hello World");
                    worksheet.Cells["C2"].PutValue("Text data type");

                    // Numeric values
                    worksheet.Cells["A3"].PutValue("Number (Integer)");
                    worksheet.Cells["B3"].PutValue(42);
                    worksheet.Cells["C3"].PutValue("Integer numeric data type");

                    worksheet.Cells["A4"].PutValue("Number (Double)");
                    worksheet.Cells["B4"].PutValue(3.14159);
                    worksheet.Cells["C4"].PutValue("Double numeric data type");

                    // Date values
                    worksheet.Cells["A5"].PutValue("Date");
                    worksheet.Cells["B5"].PutValue(DateTime.Now);
                    worksheet.Cells["C5"].PutValue("Date/time data type");

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with different cell data types created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with different cell data types: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with cells containing formulas using Openize.OpenXML-SDK.
        /// Demonstrates how to create and use different types of Excel formulas.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Cell' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "CellFormulas.xlsx").
        /// </param>
        public void CreateCellsWithFormulas(string documentDirectory = docsDirectory, string filename = "CellFormulas.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];

                    // Add headers
                    worksheet.Cells["A1"].PutValue("Formula Type");
                    worksheet.Cells["B1"].PutValue("Formula");
                    worksheet.Cells["C1"].PutValue("Result");

                    // Add data for formulas
                    worksheet.Cells["A3"].PutValue("Value 1:");
                    worksheet.Cells["B3"].PutValue(10);

                    worksheet.Cells["A4"].PutValue("Value 2:");
                    worksheet.Cells["B4"].PutValue(20);

                    worksheet.Cells["A5"].PutValue("Value 3:");
                    worksheet.Cells["B5"].PutValue(30);

                    // Add SUM formula
                    worksheet.Cells["A7"].PutValue("SUM");
                    worksheet.Cells["B7"].PutValue("=SUM(B3:B5)");
                    worksheet.Cells["C7"].PutFormula("SUM(B3:B5)");

                    // Add AVERAGE formula
                    worksheet.Cells["A8"].PutValue("AVERAGE");
                    worksheet.Cells["B8"].PutValue("=AVERAGE(B3:B5)");
                    worksheet.Cells["C8"].PutFormula("AVERAGE(B3:B5)");

                    // Add MIN formula
                    worksheet.Cells["A9"].PutValue("MIN");
                    worksheet.Cells["B9"].PutValue("=MIN(B3:B5)");
                    worksheet.Cells["C9"].PutFormula("MIN(B3:B5)");

                    // Add MAX formula
                    worksheet.Cells["A10"].PutValue("MAX");
                    worksheet.Cells["B10"].PutValue("=MAX(B3:B5)");
                    worksheet.Cells["C10"].PutFormula("MAX(B3:B5)");

                    // Add COUNT formula
                    worksheet.Cells["A11"].PutValue("COUNT");
                    worksheet.Cells["B11"].PutValue("=COUNT(B3:B5)");
                    worksheet.Cells["C11"].PutFormula("COUNT(B3:B5)");

                    // Add a more complex formula
                    worksheet.Cells["A13"].PutValue("Complex Formula");
                    worksheet.Cells["B13"].PutValue("=IF(C7>50,\"High\",\"Low\")");
                    worksheet.Cells["C13"].PutFormula("IF(C7>50,\"High\",\"Low\")");

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with cell formulas created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with cell formulas: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with hyperlinked cells using Openize.OpenXML-SDK.
        /// Demonstrates how to create cells with hyperlinks to both websites and email addresses.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Cell' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "CellHyperlinks.xlsx").
        /// </param>
        public void CreateCellsWithHyperlinks(string documentDirectory = docsDirectory, string filename = "CellHyperlinks.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];

                    // Add headers
                    worksheet.Cells["A1"].PutValue("Hyperlink Type");
                    worksheet.Cells["B1"].PutValue("Display Text");
                    worksheet.Cells["C1"].PutValue("Target");

                    // Website hyperlink
                    worksheet.Cells["A2"].PutValue("Website");
                    worksheet.Cells["B2"].PutValue("Visit Openize Website");
                    worksheet.Cells["C2"].PutValue("https://github.com/openize-com/openize-open-xml-sdk-net");

                    // Set the hyperlink
                    worksheet.Cells["B2"].SetHyperlink("https://github.com/openize-com/openize-open-xml-sdk-net", "Click to visit Openize OpenXML SDK GitHub repository");

                    // Email hyperlink
                    worksheet.Cells["A3"].PutValue("Email");
                    worksheet.Cells["B3"].PutValue("Contact Support");
                    worksheet.Cells["C3"].PutValue("mailto:support@example.com");

                    // Set the hyperlink
                    worksheet.Cells["B3"].SetHyperlink("mailto:support@example.com", "Click to send email to support");

                    // Another website hyperlink
                    worksheet.Cells["A4"].PutValue("Website");
                    worksheet.Cells["B4"].PutValue("Microsoft Documentation");
                    worksheet.Cells["C4"].PutValue("https://learn.microsoft.com/");

                    // Set the hyperlink
                    worksheet.Cells["B4"].SetHyperlink("https://learn.microsoft.com/", "Click to visit Microsoft Documentation");

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with hyperlinked cells created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with hyperlinked cells: {ex.Message}");
                throw;
            }
        }
    }
}