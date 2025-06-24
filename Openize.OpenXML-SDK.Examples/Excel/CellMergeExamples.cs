using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for merging cells in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class CellMergeExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/CellMerge";

        /// <summary>
        /// Initializes a new instance of the <see cref="CellMergeExamples"/> class.
        /// Prepares the directory 'Documents/Excel/CellMerge' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public CellMergeExamples()
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
        /// Creates a new Excel workbook with merged cells using Openize.OpenXML-SDK.
        /// Demonstrates various cell merging scenarios.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/CellMerge' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "MergedCells.xlsx").
        /// </param>
        public void CreateMergedCells(string documentDirectory = docsDirectory, string filename = "MergedCells.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                using (var workbook = new Workbook())
                {
                    var firstSheet = workbook.Worksheets[0];

                    // Example 1: Basic merge from the gist (A1 to C1)
                    firstSheet.MergeCells("A1", "C1"); // Merge cells from A1 to C1

                    // Add value to the top-left cell of the merged area
                    var topLeftCell = firstSheet.Cells["A1"];
                    topLeftCell.PutValue("This is a merged cell");
                    Console.WriteLine("Created merged cell A1:C1 with text");

                    // Example 2: Create a title header (merge A3 to E3)
                    firstSheet.MergeCells("A3", "E3");
                    firstSheet.Cells["A3"].PutValue("Employee Report - Quarter 1");
                    Console.WriteLine("Created title header in merged cells A3:E3");

                    // Example 3: Create column headers (not merged)
                    firstSheet.Cells["A5"].PutValue("Employee ID");
                    firstSheet.Cells["B5"].PutValue("Name");
                    firstSheet.Cells["C5"].PutValue("Department");
                    firstSheet.Cells["D5"].PutValue("Salary");
                    firstSheet.Cells["E5"].PutValue("Status");

                    // Example 4: Merge cells for employee names (B6 to C6)
                    firstSheet.MergeCells("B6", "C6");
                    firstSheet.Cells["A6"].PutValue("EMP001");
                    firstSheet.Cells["B6"].PutValue("John Smith");
                    firstSheet.Cells["D6"].PutValue(65000);
                    firstSheet.Cells["E6"].PutValue("Active");
                    Console.WriteLine("Created merged cell for employee name B6:C6");

                    // Example 5: Another merged name cell
                    firstSheet.MergeCells("B7", "C7");
                    firstSheet.Cells["A7"].PutValue("EMP002");
                    firstSheet.Cells["B7"].PutValue("Jane Doe");
                    firstSheet.Cells["D7"].PutValue(70000);
                    firstSheet.Cells["E7"].PutValue("Active");

                    // Example 6: Merge multiple rows for a note (A9 to E10)
                    firstSheet.MergeCells("A9", "E10");
                    firstSheet.Cells["A9"].PutValue("Note: This report shows employee data with merged cells for better presentation. Merged cells can span multiple rows and columns.");
                    Console.WriteLine("Created large merged area A9:E10 for notes");

                    // Example 7: Small merge for subtotal (D12 to E12)
                    firstSheet.MergeCells("D12", "E12");
                    firstSheet.Cells["A12"].PutValue("Total Employees:");
                    firstSheet.Cells["B12"].PutValue("2");
                    firstSheet.Cells["D12"].PutValue("Total Salary: $135,000");
                    Console.WriteLine("Created merged cell for total D12:E12");

                    workbook.Save(filePath);

                    Console.WriteLine("Cells merged successfully and workbook saved.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating merged cells: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Reads merged cell information from an existing Excel workbook.
        /// Demonstrates how to identify and work with merged cells.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook is located (default is the 'Documents/Excel/CellMerge' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file to read (default is "MergedCells.xlsx").
        /// </param>
        public void ReadMergedCells(string documentDirectory = docsDirectory, string filename = "MergedCells.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Check if the file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine("Creating the workbook first...");
                    CreateMergedCells(documentDirectory, filename);
                }

                using (var workbook = new Workbook(filePath))
                {
                    var worksheet = workbook.Worksheets[0];

                    Console.WriteLine($"Reading merged cells from worksheet: {worksheet.Name}");
                    Console.WriteLine("=====================================");

                    // Read some specific merged cells
                    Console.WriteLine("Merged Cell Contents:");
                    Console.WriteLine("--------------------");

                    Console.WriteLine($"A1 (merged A1:C1): {worksheet.Cells["A1"].GetValue()}");
                    Console.WriteLine($"A3 (merged A3:E3): {worksheet.Cells["A3"].GetValue()}");
                    Console.WriteLine($"B6 (merged B6:C6): {worksheet.Cells["B6"].GetValue()}");
                    Console.WriteLine($"B7 (merged B7:C7): {worksheet.Cells["B7"].GetValue()}");
                    Console.WriteLine($"A9 (merged A9:E10): {worksheet.Cells["A9"].GetValue()}");
                    Console.WriteLine($"D12 (merged D12:E12): {worksheet.Cells["D12"].GetValue()}");

                    // Try to read from cells that are part of merged ranges but not the top-left
                    Console.WriteLine("\nReading from non-top-left cells in merged ranges:");
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine($"B1 (part of A1:C1 merge): {worksheet.Cells["B1"].GetValue() ?? "null/empty"}");
                    Console.WriteLine($"C1 (part of A1:C1 merge): {worksheet.Cells["C1"].GetValue() ?? "null/empty"}");
                    Console.WriteLine($"C6 (part of B6:C6 merge): {worksheet.Cells["C6"].GetValue() ?? "null/empty"}");

                    Console.WriteLine("\nNote: Only the top-left cell of a merged range contains the value.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading merged cells: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Updates content in merged cells of an existing Excel workbook.
        /// Demonstrates modifying merged cell values.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook is located and where the modified workbook will be saved.
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file to modify (default is "MergedCells.xlsx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Excel workbook (default is "UpdatedMergedCells.xlsx").
        /// </param>
        public void UpdateMergedCells(string documentDirectory = docsDirectory,
            string filename = "MergedCells.xlsx", string filenameModified = "UpdatedMergedCells.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";
                string modifiedFilePath = $"{documentDirectory}/{filenameModified}";

                // Check if the file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine("Creating the workbook first...");
                    CreateMergedCells(documentDirectory, filename);
                }

                using (var workbook = new Workbook(filePath))
                {
                    var worksheet = workbook.Worksheets[0];

                    Console.WriteLine($"Updating merged cells in worksheet: {worksheet.Name}");

                    // Update the main title
                    worksheet.Cells["A1"].PutValue("Updated: This is a modified merged cell");
                    Console.WriteLine("Updated main title in A1:C1");

                    // Update the report title
                    worksheet.Cells["A3"].PutValue("Employee Report - Quarter 2 (Updated)");
                    Console.WriteLine("Updated report title in A3:E3");

                    // Update employee names
                    worksheet.Cells["B6"].PutValue("John Smith (Senior)");
                    worksheet.Cells["B7"].PutValue("Jane Doe (Manager)");
                    Console.WriteLine("Updated employee names in merged cells");

                    // Update the note
                    worksheet.Cells["A9"].PutValue("Updated Note: This report has been modified to show Quarter 2 data. All merged cells have been updated with new information. Last updated: " + DateTime.Now.ToString("yyyy-MM-dd"));
                    Console.WriteLine("Updated note in merged area A9:E10");

                    // Update the total
                    worksheet.Cells["D12"].PutValue("Updated Total Salary: $145,000");
                    Console.WriteLine("Updated total in merged cells D12:E12");

                    // Add a new merged cell for timestamp
                    worksheet.MergeCells("A14", "E14");
                    worksheet.Cells["A14"].PutValue($"Last Modified: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    Console.WriteLine("Added new merged timestamp cell A14:E14");

                    workbook.Save(modifiedFilePath);
                    Console.WriteLine($"\nUpdated workbook saved as {filenameModified}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating merged cells: {ex.Message}");
                throw;
            }
        }
    }
}