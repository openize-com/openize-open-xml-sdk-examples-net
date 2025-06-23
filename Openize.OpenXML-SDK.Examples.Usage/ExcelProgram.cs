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
                        RunCellValueExamples();
                        break;

                    case "2":
                        RunCellMergeExamples();
                        break;

                    case "3":
                        RunRowColumnExamples();
                        break;

                    case "4":
                        RunFreezePanesExamples();
                        break;

                    case "5":
                        RunRangeExamples();
                        break;

                    case "6":
                        RunDocumentPropertiesExamples();
                        break;

                    case "7":
                        RunColumnInsertionExamples();
                        break;

                    case "8":
                        RunRowInsertionExamples();
                        break;

                    case "9":
                        RunAddWorksheetExamples();
                        break;

                    case "10":
                        RunFormulaExamples();
                        break;

                    case "11":
                        RunHiddenSheetsExamples();
                        break;

                    case "12":
                        RunCellStylingExamples();
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
            Console.WriteLine("1. Cell Value Examples");
            Console.WriteLine("2. Cell Merge Examples");
            Console.WriteLine("3. Row Column Examples");
            Console.WriteLine("4. Freeze Panes Examples");
            Console.WriteLine("5. Range Examples");
            Console.WriteLine("6. Document Properties Examples");
            Console.WriteLine("7. Column Insertion Examples");
            Console.WriteLine("8. Row Insertion Examples");
            Console.WriteLine("9. Add Worksheet Examples");
            Console.WriteLine("10. Formula Examples");
            Console.WriteLine("11. Hidden Sheets Examples");
            Console.WriteLine("12. Cell Styling Examples");
            Console.WriteLine("0. Back to Main Menu");
            Console.Write("\nEnter your choice: ");
        }

        private static void RunCellValueExamples()
        {
            Console.Clear();
            Console.WriteLine("Cell Value Examples");
            Console.WriteLine("==================");

            var cellValueExamples = new CellValueExamples();

            Console.WriteLine("\n1. Creating cells with basic values...");
            cellValueExamples.CreateCellValues();

            Console.WriteLine("\n2. Reading cell values from workbook...");
            cellValueExamples.ReadCellValues();

            Console.WriteLine("\n3. Updating cell values in workbook...");
            cellValueExamples.UpdateCellValues();
        }

        private static void RunCellMergeExamples()
        {
            Console.Clear();
            Console.WriteLine("Cell Merge Examples");
            Console.WriteLine("==================");

            var cellMergeExamples = new CellMergeExamples();

            Console.WriteLine("\n1. Creating merged cells...");
            cellMergeExamples.CreateMergedCells();

            Console.WriteLine("\n2. Reading merged cells...");
            cellMergeExamples.ReadMergedCells();

            Console.WriteLine("\n3. Updating merged cells...");
            cellMergeExamples.UpdateMergedCells();
        }

        private static void RunRowColumnExamples()
        {
            Console.Clear();
            Console.WriteLine("Row Column Examples");
            Console.WriteLine("==================");

            var rowColumnExamples = new RowColumnExamples();

            Console.WriteLine("\n1. Creating row and column sizing...");
            rowColumnExamples.CreateRowColumnSizing();
        }

        private static void RunFreezePanesExamples()
        {
            Console.Clear();
            Console.WriteLine("Freeze Panes Examples");
            Console.WriteLine("====================");

            var freezePanesExamples = new FreezePanesExamples();

            Console.WriteLine("\n1. Creating freeze panes...");
            freezePanesExamples.CreateFreezePanes();
        }

        private static void RunRangeExamples()
        {
            Console.Clear();
            Console.WriteLine("Range Examples");
            Console.WriteLine("=============");

            var rangeExamples = new RangeExamples();

            Console.WriteLine("\n1. Creating range example...");
            rangeExamples.CreateRangeExample();
        }

        private static void RunDocumentPropertiesExamples()
        {
            Console.Clear();
            Console.WriteLine("Document Properties Examples");
            Console.WriteLine("===========================");

            var documentPropertiesExamples = new DocumentPropertiesExamples();

            Console.WriteLine("\n1. Creating document properties...");
            documentPropertiesExamples.CreateDocumentProperties();
        }

        private static void RunColumnInsertionExamples()
        {
            Console.Clear();
            Console.WriteLine("Column Insertion Examples");
            Console.WriteLine("=========================");

            var columnInsertionExamples = new ColumnInsertionExamples();

            Console.WriteLine("\n1. Creating column insertion...");
            columnInsertionExamples.CreateColumnInsertion();
        }

        private static void RunRowInsertionExamples()
        {
            Console.Clear();
            Console.WriteLine("Row Insertion Examples");
            Console.WriteLine("======================");

            var rowInsertionExamples = new RowInsertionExamples();

            Console.WriteLine("\n1. Creating row insertion...");
            rowInsertionExamples.CreateRowInsertion();
        }

        private static void RunAddWorksheetExamples()
        {
            Console.Clear();
            Console.WriteLine("Add Worksheet Examples");
            Console.WriteLine("======================");

            var addWorksheetExamples = new AddWorksheetExamples();

            Console.WriteLine("\n1. Creating add worksheet...");
            addWorksheetExamples.CreateAddWorksheet();
        }

        private static void RunFormulaExamples()
        {
            Console.Clear();
            Console.WriteLine("Formula Examples");
            Console.WriteLine("================");

            var formulaExamples = new FormulaExamples();

            Console.WriteLine("\n1. Creating formula example...");
            formulaExamples.CreateFormulaExample();
        }

        private static void RunHiddenSheetsExamples()
        {
            Console.Clear();
            Console.WriteLine("Hidden Sheets Examples");
            Console.WriteLine("======================");

            var hiddenSheetsExamples = new HiddenSheetsExamples();

            Console.WriteLine("\n1. Creating hidden sheets...");
            hiddenSheetsExamples.CreateHiddenSheets();

            Console.WriteLine("\n2. Getting hidden sheets information...");
            hiddenSheetsExamples.GetHiddenSheets();
        }

        private static void RunCellStylingExamples()
        {
            Console.Clear();
            Console.WriteLine("Cell Styling Examples");
            Console.WriteLine("====================");

            var cellStylingExamples = new CellStylingExamples();

            Console.WriteLine("\n1. Creating cell styling...");
            cellStylingExamples.CreateCellStyling();
        }
    }
}