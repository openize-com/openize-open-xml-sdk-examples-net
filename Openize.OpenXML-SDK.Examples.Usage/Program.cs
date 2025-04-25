using System;
using Openize.OpenXML_SDK.Examples.Excel;

namespace Openize.OpenXML_SDK.Examples.Usage
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Openize.OpenXML-SDK Examples");
            Console.WriteLine("============================");
            Console.WriteLine();

            bool exit = false;

            while (!exit)
            {
                DisplayMenu();
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        RunWorkbookExamples();
                        break;

                    case "2":
                        RunWorksheetExamples();
                        break;

                    case "3":
                        RunCellExamples();
                        break;

                    case "0":
                        exit = true;
                        break;

                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }

                if (!exit)
                {
                    Console.WriteLine("\nPress any key to return to the menu...");
                    Console.ReadKey();
                    Console.Clear();
                }
            }

            Console.WriteLine("\nThank you for exploring the Openize.OpenXML-SDK Examples!");
        }

        static void DisplayMenu()
        {
            Console.WriteLine("Choose an example to run:");
            Console.WriteLine("1. Workbook Examples");
            Console.WriteLine("2. Worksheet Examples");
            Console.WriteLine("3. Cell Examples");
            Console.WriteLine("0. Exit");
            Console.Write("\nEnter your choice: ");
        }

        static void RunWorkbookExamples()
        {
            Console.Clear();
            Console.WriteLine("Workbook Examples");
            Console.WriteLine("================");

            var workbookExamples = new WorkbookExamples();

            Console.WriteLine("\n1. Creating an empty workbook...");
            workbookExamples.CreateEmptyWorkbook();

            Console.WriteLine("\n2. Creating a workbook with multiple sheets...");
            workbookExamples.CreateWorkbookWithMultipleSheets();

            Console.WriteLine("\n3. Creating a workbook with properties...");
            workbookExamples.CreateWorkbookWithProperties();

            Console.WriteLine("\n4. Opening an existing workbook...");
            workbookExamples.OpenExistingWorkbook();
        }

        static void RunWorksheetExamples()
        {
            Console.Clear();
            Console.WriteLine("Worksheet Examples");
            Console.WriteLine("=================");

            var worksheetExamples = new WorksheetExamples();

            Console.WriteLine("\n1. Creating a renamed worksheet...");
            worksheetExamples.CreateRenamedWorksheet();

            Console.WriteLine("\n2. Creating a protected worksheet...");
            worksheetExamples.CreateProtectedWorksheet();

            Console.WriteLine("\n3. Modifying column width and row height...");
            worksheetExamples.ModifyColumnWidthAndRowHeight();

            Console.WriteLine("\n4. Working with hidden columns and rows...");
            worksheetExamples.HideColumnsAndRows();

            Console.WriteLine("\n5. Working with merged cells...");
            worksheetExamples.MergeCells();
        }

        static void RunCellExamples()
        {
            Console.Clear();
            Console.WriteLine("Cell Examples");
            Console.WriteLine("=============");

            var cellExamples = new CellExamples();

            Console.WriteLine("\n1. Creating cells with different data types...");
            cellExamples.CreateCellsWithDifferentDataTypes();

            Console.WriteLine("\n2. Creating cells with formulas...");
            cellExamples.CreateCellsWithFormulas();

            Console.WriteLine("\n3. Creating cells with hyperlinks...");
            cellExamples.CreateCellsWithHyperlinks();
        }
    }
}