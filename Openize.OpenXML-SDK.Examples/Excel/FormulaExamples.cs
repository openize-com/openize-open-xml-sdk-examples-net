using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for working with formulas in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class FormulaExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/Formula";

        /// <summary>
        /// Initializes a new instance of the <see cref="FormulaExamples"/> class.
        /// Prepares the directory 'Documents/Excel/Formula' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public FormulaExamples()
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
        /// Creates a new Excel workbook and demonstrates formula usage using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/Formula' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "FormulaExample.xlsx").
        /// </param>
        public void CreateFormulaExample(string documentDirectory = docsDirectory, string filename = "FormulaExample.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                using (var wb = new Workbook())
                {
                    // Accessing the first worksheet in the workbook.
                    var firstSheet = wb.Worksheets[0];

                    // Create a random number generator.
                    Random rand = new Random();

                    // Loop through the first 10 rows in column A.
                    for (int i = 1; i <= 10; i++)
                    {
                        // Construct a cell reference based on the current row.
                        string cellReference = $"A{i}";

                        // Generate a random number between 1 and 100.
                        double randomValue = rand.Next(1, 100);

                        // Assign the random number to the cell.
                        firstSheet.Cells[cellReference].PutValue(randomValue);
                    }

                    // After populating the first 10 cells with random numbers,
                    // we will use cell A11 to sum the values from A1 to A10.
                    firstSheet.Cells["A11"].PutFormula("SUM(A1:A10)");

                    // Save the changes made to the workbook.
                    wb.Save(filePath);

                    Console.WriteLine("Formula example created and workbook saved successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating formula example: {ex.Message}");
                throw;
            }
        }
    }
}