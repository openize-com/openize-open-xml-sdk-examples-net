using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for working with worksheet properties and display settings
    /// using the Openize.OpenXML-SDK library.
    /// </summary>
    public class WorksheetPropertiesExamples
    {
        private const string docsDirectory = "../../../Documents/WorksheetProperties";

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetPropertiesExamples"/> class.
        /// Prepares the directory 'Documents/WorksheetProperties' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public WorksheetPropertiesExamples()
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
        /// Creates a workbook demonstrating worksheet properties if the extensions are available.
        /// </summary>
        /// <param name="documentDirectory">The directory to save the workbook in.</param>
        /// <param name="filename">The filename for the workbook.</param>
        public void DemonstrateWorksheetProperties(string documentDirectory = docsDirectory, string filename = "WorksheetProperties.xlsx")
        {
            try
            {
                using (var workbook = new Workbook())
                {
                    var worksheet = workbook.Worksheets[0];
                    worksheet.Name = "Properties Demo";

                    // Add basic content
                    worksheet.Cells["A1"].PutValue("Worksheet Properties Example");
                    worksheet.Cells["A2"].PutValue("This worksheet demonstrates various properties");

                    // Add some sample data
                    for (int row = 4; row <= 10; row++)
                    {
                        for (int col = 0; col < 5; col++)
                        {
                            string cellRef = $"{(char)('A' + col)}{row}";
                            worksheet.Cells[cellRef].PutValue($"Cell {cellRef}");
                        }
                    }

                    // Try to use extensions if available
                    // Comment these out if they're not available in your version
                    
                    worksheet.SetZoom(150);
                    worksheet.ShowGridlines(false);
                    worksheet.ShowRowColumnHeaders(false);
                    worksheet.ShowZeroValues(false);
                    worksheet.SetDefaultColumnWidth(15);
                    worksheet.SetDefaultRowHeight(25);
                    

                    // Alternative: Add a note about the features
                    worksheet.Cells["A12"].PutValue("Note: Worksheet properties extension methods would be demonstrated here");
                    worksheet.Cells["A13"].PutValue("Features include: Zoom, Gridlines, Headers, Zero Values, Default Dimensions");

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");

                    Console.WriteLine("Note: Worksheet properties extensions demonstration requires the latest SDK version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error demonstrating worksheet properties: {ex.Message}");
                throw;
            }
        }
    }
}