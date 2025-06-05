using System;
using Openize.OpenXML_SDK.Examples.Excel;
using Openize.Slides.Examples;

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

                    case "4":
                        RunSlideExamples();
                        break;

                    case "5":
                        RunSlideTextExamples();
                        break;

                    case "6":
                        RunSlideImageExamples();
                        break;

                    case "7":
                        RunSlideStyledListExamples();
                        break;

                    case "8":
                        RunSlideTableExamples();
                        break;

                    case "9":
                        RunSlideCommentExamples();
                        break;

                    case "10":
                        RunSlideCommentAuthorExamples();
                        break;

                    case "11":
                        RunSlideNotesExamples();
                        break;

                    case "12":
                        RunSlideRectangleExamples();
                        break;

                    case "13":
                        RunSlideCircleExamples();
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
            Console.WriteLine("4. Slide Examples");
            Console.WriteLine("5. Slide Text Examples");
            Console.WriteLine("6. Slide Image Examples");
            Console.WriteLine("7. Slide Styled List Examples");
            Console.WriteLine("8. Slide Table Examples");
            Console.WriteLine("9. Slide Comment Examples");
            Console.WriteLine("10. Slide Comment Author Examples");
            Console.WriteLine("11. Slide Notes Examples");
            Console.WriteLine("12. Slide Rectangle Examples");
            Console.WriteLine("13. Slide Circle Examples");
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
        static void RunSlideExamples()
        {
            Console.Clear();
            Console.WriteLine("Slide Examples");
            Console.WriteLine("==============");

            var example = new SlideExamples();

            Console.WriteLine("\n1. SetDimensionsOfSlides...");
            example.SetDimensionsOfSlides();

            Console.WriteLine("\n2. CreateNewSlideInNewPresentation...");
            example.CreateNewSlideInNewPresentation();

            Console.WriteLine("\n3. CreateNewSlideInExistingPresentation...");
            example.CreateNewSlideInExistingPresentation(filename: "test.pptx");

            Console.WriteLine("\n4. RemoveSlideInAnExistingPresentation...");
            example.RemoveSlideInAnExistingPresentation(filename: "test.pptx");

            Console.WriteLine("\n5. AddBackgroundColorToAnExistingSlide...");
            example.AddBackgroundColorToAnExistingSlide(filename: "sample.pptx");
        }

        static void RunSlideTextExamples()
        {
            Console.Clear();
            Console.WriteLine("Text Examples");
            Console.WriteLine("=============");

            var example = new TextExamples();

            Console.WriteLine("\n1. CreateTextShapeInNewSlide...");
            example.CreateNewTextShapeInNewSlide();

            Console.WriteLine("\n2. AddTextShapeToExistingSlide...");
            example.AddNewTextShapeExistingSlide(filename: "sample.pptx");
        }

        static void RunSlideImageExamples()
        {
            Console.Clear();
            Console.WriteLine("Image Examples");
            Console.WriteLine("==============");

            var example = new ImageExamples();

            Console.WriteLine("\n1. AddImageInASlide...");
            example.AddImageInASlide(imagename: "sample.jpg");

            Console.WriteLine("\n2. UpdateImageInExistingSlide...");
            example.UpdateImageInExistingSlide(filename: "sample.pptx", xAxis: 300.0, yAxis: 200.0);
        }

        static void RunSlideStyledListExamples()
        {
            Console.Clear();
            Console.WriteLine("Styled List Examples");
            Console.WriteLine("====================");

            var example = new StyledListExamples();

            Console.WriteLine("\n1. CreateBulletedListInASlide...");
            example.CreateBulletedListInASlide();

            Console.WriteLine("\n2. AddListItemsInAnExistingList...");
            example.AddListItemsInAnExistingList(filename: "test.pptx");

            Console.WriteLine("\n3. RemoveListItemsInAnExistingList...");
            example.RemoveListItemsInAnExistingList(filename: "test.pptx");
        }

        static void RunSlideTableExamples()
        {
            Console.Clear();
            Console.WriteLine("Table Examples");
            Console.WriteLine("==============");

            var example = new TableExamples();

            Console.WriteLine("\n1. CreateSimpleTableInASlide...");
            example.CreateSimpleTableInASlide(filename: "sample.pptx");

            Console.WriteLine("\n2. CreateTableWithTableStylingsInASlide...");
            example.CreateTableWithTableStylingsInASlide(filename: "sample.pptx");

            Console.WriteLine("\n3. CreateTableWithRowStylingsInASlide...");
            example.CreateTableWithRowStylingsInASlide(filename: "sample.pptx");

            Console.WriteLine("\n4. CreateTableWithCellStylingsInASlide...");
            example.CreateTableWithCellStylingsInASlide(filename: "sample.pptx");

            Console.WriteLine("\n5. CreateTableWithThemeInASlide...");
            example.CreateTableWithThemeInASlide(filename: "sample.pptx");

            Console.WriteLine("\n6. AddRowInAnExistingTableInASlide...");
            example.AddRowInAnExistingTableInASlide(filename: "sample.pptx");

            Console.WriteLine("\n7. AddColumnWithCellValuesInAnExistingTableInASlide...");
            example.AddColumnWithCellValuesInAnExistingTableInASlide(filename: "sample.pptx");
        }

        static void RunSlideCommentExamples()
        {
            Console.Clear();
            Console.WriteLine("Comment Examples");
            Console.WriteLine("================");

            var example = new CommentExamples();

            Console.WriteLine("\n1. CreateCommentInASlide...");
            example.CreateCommentInASlide(filename: "sample.pptx");

            Console.WriteLine("\n2. RemoveACommentFromASlide...");
            example.RemoveACommentFromASlide(filename: "sample.pptx");

            Console.WriteLine("\n3. AddCommentWithExistingCommentAuthor...");
            example.AddCommentWithExistingCommentAuthor(filename: "sample.pptx");
        }

        static void RunSlideCommentAuthorExamples()
        {
            Console.Clear();
            Console.WriteLine("Comment Author Examples");
            Console.WriteLine("========================");

            var example = new CommentAuthorExamples();

            Console.WriteLine("\n1. AddCommentAuthor...");
            example.AddCommentAuthor(filename: "sample.pptx");

            Console.WriteLine("\n2. RemoveCommentAuthor...");
            example.RemoveCommentAuthor(filename: "sample.pptx");
        }

        static void RunSlideNotesExamples()
        {
            Console.Clear();
            Console.WriteLine("Notes Examples");
            Console.WriteLine("==============");

            var example = new NotesExamples();

            Console.WriteLine("\n1. CreateNotesInASlide...");
            example.CreateNotesInASlide(filename: "sample.pptx");

            Console.WriteLine("\n2. RemoveNotesFromASlide...");
            example.RemoveNotesFromASlide(filename: "sample.pptx");

            Console.WriteLine("\n3. ExportNotesToTextFile...");
            example.ExportNotesToTextFile(filename: "sample.pptx");
        }

        static void RunSlideRectangleExamples()
        {
            Console.Clear();
            Console.WriteLine("Rectangle Examples");
            Console.WriteLine("==================");

            var example = new RectangleExamples();

            Console.WriteLine("\n1. DrawNewRectangleShapeInNewSlide...");
            example.DrawNewRectangleShapeInNewSlide(filename: "sample.pptx");

            Console.WriteLine("\n2. RemoveRectangleShapeExistingSlide...");
            example.RemoveRectangleShapeExistingSlide(filename: "sample.pptx");

            Console.WriteLine("\n3. SetBackgroundColorOfRectangle...");
            example.SetBackgroundColorOfRectangle(filename: "sample.pptx");

            Console.WriteLine("\n4. DrawNewRectangleShapeWithAnimation...");
            example.DrawNewRectangleShapeWithAnimation(filename: "sample.pptx");
        }

        static void RunSlideCircleExamples()
        {
            Console.Clear();
            Console.WriteLine("Circle Examples");
            Console.WriteLine("===============");

            var example = new CircleExamples();

            Console.WriteLine("\n1. DrawCircleShapeInNewSlide...");
            example.DrawCircleShapeInNewSlide(filename: "sample.pptx");

            Console.WriteLine("\n2. SetBackgroundColorOfCircle...");
            example.SetBackgroundColorOfCircle(filename: "sample.pptx");

            Console.WriteLine("\n3. RemoveCircleShapeExistingSlide...");
            example.RemoveCircleShapeExistingSlide(filename: "sample.pptx");

            Console.WriteLine("\n4. DrawCircleShapeWithAnimation...");
            example.DrawCircleShapeWithAnimation(filename: "sample.pptx");
        }



    }
}