using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Excel worksheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Worksheet at the root of your project.
    /// // Check reference for more options and details.
    /// WorksheetExamples worksheetExamples = new WorksheetExamples();
    /// // Creates a workbook with a renamed worksheet and saves it to the specified directory.
    /// // Check reference for more options and details.
    /// worksheetExamples.CreateRenamedWorksheet();
    /// // Creates a workbook with worksheet protection and saves it to the specified directory.
    /// // Check reference for more options and details.
    /// worksheetExamples.CreateProtectedWorksheet();
    /// // Modifies column width and row height in a worksheet.
    /// // Check reference for more options and details.
    /// worksheetExamples.ModifyColumnWidthAndRowHeight();
    /// </code>
    /// </example>
    public class WorksheetExamples
    {
        private const string docsDirectory = "../../../Documents/Worksheet";

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetExamples"/> class.
        /// Prepares the directory 'Documents/Worksheet' for storing or loading Excel workbooks
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public WorksheetExamples()
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
        /// Creates a new Excel workbook with a renamed worksheet using Openize.OpenXML-SDK.
        /// Renames the default worksheet to "Data" and saves the workbook.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Worksheet' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "RenamedWorksheet.xlsx").
        /// </param>
        public void CreateRenamedWorksheet(string documentDirectory = docsDirectory, string filename = "RenamedWorksheet.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    // Rename the default worksheet
                    workbook.Worksheets[0].Name = "Data";

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with renamed worksheet created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with renamed worksheet: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with a protected worksheet using Openize.OpenXML-SDK.
        /// Adds protection to the default worksheet with a password and saves the workbook.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Worksheet' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "ProtectedWorksheet.xlsx").
        /// </param>
        public void CreateProtectedWorksheet(string documentDirectory = docsDirectory, string filename = "ProtectedWorksheet.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    // Rename the default worksheet
                    workbook.Worksheets[0].Name = "Protected";

                    // Add some data to demonstrate that it can't be modified after protection
                    workbook.Worksheets[0].Cells["A1"].PutValue("This worksheet is protected");
                    workbook.Worksheets[0].Cells["A2"].PutValue("You cannot modify cells without the password");

                    // Protect the worksheet with a password
                    workbook.Worksheets[0].ProtectSheet("password123");

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with protected worksheet created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with protected worksheet: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook and modifies column widths and row heights using Openize.OpenXML-SDK.
        /// Demonstrates how to set custom widths and heights for specific columns and rows.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Worksheet' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "ColumnWidthRowHeight.xlsx").
        /// </param>
        public void ModifyColumnWidthAndRowHeight(string documentDirectory = docsDirectory, string filename = "ColumnWidthRowHeight.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];

                    // Add headers
                    worksheet.Cells["A1"].PutValue("Column A (Normal Width)");
                    worksheet.Cells["B1"].PutValue("Column B (Wide)");
                    worksheet.Cells["C1"].PutValue("Column C (Very Wide)");
                    worksheet.Cells["D1"].PutValue("Column D (Narrow)");

                    // Add data
                    worksheet.Cells["A2"].PutValue("Normal row height");
                    worksheet.Cells["A3"].PutValue("Tall row height");
                    worksheet.Cells["A4"].PutValue("Very tall row height");

                    // Set column widths
                    worksheet.SetColumnWidth("B", 20); // Wide column
                    worksheet.SetColumnWidth("C", 30); // Very wide column
                    worksheet.SetColumnWidth("D", 5);  // Narrow column

                    // Set row heights
                    worksheet.SetRowHeight(3, 30);  // Tall row
                    worksheet.SetRowHeight(4, 50);  // Very tall row

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with custom column widths and row heights created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error modifying column widths and row heights: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook and demonstrates how to hide and unhide columns and rows.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Worksheet' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "HiddenColumnsRows.xlsx").
        /// </param>
        public void HideColumnsAndRows(string documentDirectory = docsDirectory, string filename = "HiddenColumnsRows.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];

                    // Add column headers
                    for (int col = 0; col < 10; col++)
                    {
                        var colLetter = (char)('A' + col);
                        worksheet.Cells[$"{colLetter}1"].PutValue($"Column {colLetter}");
                    }

                    // Add row headers
                    for (int row = 1; row <= 10; row++)
                    {
                        worksheet.Cells[$"A{row}"].PutValue($"Row {row}");
                    }

                    // Add data to all cells
                    for (int col = 1; col < 10; col++)
                    {
                        var colLetter = (char)('A' + col);
                        for (int row = 2; row <= 10; row++)
                        {
                            worksheet.Cells[$"{colLetter}{row}"].PutValue($"{colLetter}{row}");
                        }
                    }

                    // Hide columns C and E
                    worksheet.HideColumn("C");
                    worksheet.HideColumn("E");

                    // Hide rows 4, 6, and 8
                    worksheet.HideRow(4);
                    worksheet.HideRow(6);
                    worksheet.HideRow(8);

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with hidden columns and rows created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with hidden columns and rows: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook and demonstrates how to merge cells.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Worksheet' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "MergedCells.xlsx").
        /// </param>
        public void MergeCells(string documentDirectory = docsDirectory, string filename = "MergedCells.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];

                    // Add a title in a merged cell
                    worksheet.Cells["A1"].PutValue("Merged Cells Example");
                    worksheet.MergeCells("A1", "E1");

                    // Add a subtitle in another merged cell
                    worksheet.Cells["A2"].PutValue("This demonstrates how to merge cells in Excel");
                    worksheet.MergeCells("A2", "E2");

                    // Add some data
                    worksheet.Cells["A4"].PutValue("Category");
                    worksheet.Cells["B4"].PutValue("January");
                    worksheet.Cells["C4"].PutValue("February");
                    worksheet.Cells["D4"].PutValue("March");
                    worksheet.Cells["E4"].PutValue("Total");

                    worksheet.Cells["A5"].PutValue("Products");
                    worksheet.Cells["B5"].PutValue(1000);
                    worksheet.Cells["C5"].PutValue(1200);
                    worksheet.Cells["D5"].PutValue(1400);
                    worksheet.Cells["E5"].PutFormula("SUM(B5:D5)");

                    worksheet.Cells["A6"].PutValue("Services");
                    worksheet.Cells["B6"].PutValue(800);
                    worksheet.Cells["C6"].PutValue(850);
                    worksheet.Cells["D6"].PutValue(900);
                    worksheet.Cells["E6"].PutFormula("SUM(B6:D6)");

                    // Merge cells for a note
                    worksheet.Cells["A8"].PutValue("Note: This is a sample report with merged cells for demonstration purposes.");
                    worksheet.MergeCells("A8", "E8");

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with merged cells created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with merged cells: {ex.Message}");
                throw;
            }
        }
    }
}