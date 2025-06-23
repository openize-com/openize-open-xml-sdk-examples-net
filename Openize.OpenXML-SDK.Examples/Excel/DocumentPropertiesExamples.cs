using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for working with document properties in Excel spreadsheets
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    public class DocumentPropertiesExamples
    {
        private const string docsDirectory = "../../../Documents/Excel/DocumentProperties";

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentPropertiesExamples"/> class.
        /// Prepares the directory 'Documents/Excel/DocumentProperties' for storing or loading Excel workbooks
        /// at the root of the project.
        /// </summary>
        public DocumentPropertiesExamples()
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
        /// Creates a new Excel workbook with document properties using Openize.OpenXML-SDK.
        /// Based on the provided example.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Excel/DocumentProperties' directory).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "DocumentProperties.xlsx").
        /// </param>
        public void CreateDocumentProperties(string documentDirectory = docsDirectory, string filename = "DocumentProperties.xlsx")
        {
            try
            {
                string filePath = $"{documentDirectory}/{filename}";

                // Create a new workbook and set cell values and document properties.
                using (var workbook = new Workbook())
                {
                    // Access the first worksheet.
                    var firstSheet = workbook.Worksheets[0];

                    // Set values for cells A1 and A2.
                    firstSheet.Cells["A1"].PutValue("Text A1");
                    firstSheet.Cells["A2"].PutValue("Text A2");

                    // Configure document properties.
                    var newProperties = new BuiltInDocumentProperties
                    {
                        Author = "Fahad Adeel",
                        Title = "Sample Workbook",
                        CreatedDate = DateTime.Now,
                        ModifiedBy = "Fahad",
                        ModifiedDate = DateTime.Now.AddHours(1),
                        Subject = "Testing Subject"
                    };

                    // Assign the new properties to the workbook.
                    workbook.BuiltinDocumentProperties = newProperties;

                    // Save the workbook to the specified path.
                    workbook.Save(filePath);

                    Console.WriteLine("Document properties set and workbook saved successfully.");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating document properties: {ex.Message}");
                throw;
            }
        }
    }
}