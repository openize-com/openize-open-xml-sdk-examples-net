using System;
using System.IO;
using Openize.Cells;

namespace Openize.OpenXML_SDK.Examples.Excel
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Excel workbooks
    /// using the <a href="https://github.com/openize-com/openize-open-xml-sdk-net">Openize.OpenXML-SDK</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Workbook at the root of your project.
    /// // Check reference for more options and details.
    /// WorkbookExamples workbookExamples = new WorkbookExamples();
    /// // Creates an empty workbook and saves it to the specified directory.
    /// // Check reference for more options and details.
    /// workbookExamples.CreateEmptyWorkbook();
    /// // Creates a workbook with multiple sheets and saves it to the specified directory.
    /// // Check reference for more options and details.
    /// workbookExamples.CreateWorkbookWithMultipleSheets();
    /// // Creates a workbook with built-in document properties and saves it to the specified directory.
    /// // Check reference for more options and details.
    /// workbookExamples.CreateWorkbookWithProperties();
    /// </code>
    /// </example>
    public class WorkbookExamples
    {
        private const string docsDirectory = "../../../Documents/Workbook";

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookExamples"/> class.
        /// Prepares the directory 'Documents/Workbook' for storing or loading Excel workbooks
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public WorkbookExamples()
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
        /// Creates a new empty Excel workbook using Openize.OpenXML-SDK.
        /// Saves the newly created workbook to the specified location.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Workbook' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "EmptyWorkbook.xlsx").
        /// </param>
        public void CreateEmptyWorkbook(string documentDirectory = docsDirectory, string filename = "EmptyWorkbook.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Empty workbook created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with multiple worksheets using Openize.OpenXML-SDK.
        /// Adds three sheets (Sheet1, Sheet2, and Sheet3) and saves the workbook.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Workbook' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "MultipleSheets.xlsx").
        /// </param>
        public void CreateWorkbookWithMultipleSheets(string documentDirectory = docsDirectory, string filename = "MultipleSheets.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    // The workbook already has one default sheet

                    // Add two more sheets
                    workbook.AddSheet("Sheet2");
                    workbook.AddSheet("Sheet3");

                    // Rename the first sheet
                    workbook.Worksheets[0].Name = "Sheet1";

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with multiple sheets created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with multiple sheets: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates a new Excel workbook with built-in document properties using Openize.OpenXML-SDK.
        /// Sets properties like Author, Title, Subject, and Created/Modified dates.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook will be saved (default is the 'Documents/Workbook' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file (default is "WorkbookWithProperties.xlsx").
        /// </param>
        public void CreateWorkbookWithProperties(string documentDirectory = docsDirectory, string filename = "WorkbookWithProperties.xlsx")
        {
            try
            {
                // Create a new workbook
                using (var workbook = new Workbook())
                {
                    // Set built-in document properties
                    workbook.BuiltinDocumentProperties = new BuiltInDocumentProperties
                    {
                        Author = "Openize SDK Example",
                        Title = "Excel Workbook with Properties",
                        Subject = "Demonstrating Document Properties",
                        CreatedDate = DateTime.Now,
                        ModifiedDate = DateTime.Now,
                        ModifiedBy = "Openize SDK Example"
                    };

                    // Save the workbook
                    workbook.Save($"{documentDirectory}/{filename}");
                    Console.WriteLine($"Workbook with properties created and saved as {filename}");
                    Console.WriteLine($"Please check directory: {Path.GetFullPath(documentDirectory)}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating workbook with properties: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Opens an existing Excel workbook and displays information about its sheets.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Excel workbook is located (default is the 'Documents/Workbook' directory at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Excel workbook file to open (default is "MultipleSheets.xlsx").
        /// </param>
        public void OpenExistingWorkbook(string documentDirectory = docsDirectory, string filename = "MultipleSheets.xlsx")
        {
            try
            {
                // Check if the file exists
                string filePath = $"{documentDirectory}/{filename}";
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    Console.WriteLine("Creating a workbook with multiple sheets first...");
                    CreateWorkbookWithMultipleSheets(documentDirectory, filename);
                }

                // Open the existing workbook
                using (var workbook = new Workbook(filePath))
                {
                    // Display information about the workbook
                    Console.WriteLine($"Opened workbook: {filename}");
                    Console.WriteLine($"Number of sheets: {workbook.Worksheets.Count}");

                    // Display sheet names
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        Console.WriteLine($"Sheet {i + 1}: {workbook.Worksheets[i].Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening workbook: {ex.Message}");
                throw;
            }
        }
    }
}