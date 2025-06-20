using System;
using Openize.OpenXML_SDK.Examples.Excel;

namespace Openize.OpenXML_SDK.Examples.Usage
{
    public static class ExcelProgram
    {
        public static void Run()
        {
            bool back = false;

            while (!back)
            {
                Console.Clear();
                DisplayExcelMenu();
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

                    case "4":
                        RunWorksheetPropertiesExamples();
                        break;


                    case "0":
                        back = true;
                        break;

                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }

                if (!back)
                {
                    Console.WriteLine("\nPress any key to return to the Excel menu...");
                    Console.ReadKey();
                }
            }
        }

        private static void DisplayExcelMenu()
        {
            Console.WriteLine("Excel Examples");
            Console.WriteLine("==============");
            Console.WriteLine("Choose an example to run:");
            Console.WriteLine("1. Workbook Examples");
            Console.WriteLine("2. Worksheet Examples");
            Console.WriteLine("3. Cell Examples");
            Console.WriteLine("4. Worksheet Properties Examples");
            Console.WriteLine("5. Data Import/Export Examples");
            Console.WriteLine("0. Back to Main Menu");
            Console.Write("\nEnter your choice: ");
        }

        private static void RunWorkbookExamples()
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

        private static void RunWorksheetExamples()
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

        private static void RunCellExamples()
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

        private static void RunWorksheetPropertiesExamples()
        {
            Console.Clear();
            Console.WriteLine("Worksheet Properties Examples");
            Console.WriteLine("============================");

            var worksheetPropertiesExamples = new WorksheetPropertiesExamples();

            Console.WriteLine("\n1. Demonstrating worksheet properties...");
            worksheetPropertiesExamples.DemonstrateWorksheetProperties();
        }
    }
}