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

                    case "14":
                        RunWordParagraphExamples();
                        break;

                    case "15":
                        RunWordParagraphAlignmentExamples();
                        break;

                    case "16":
                        RunWordParagraphIndentExamples();
                        break;

                    case "17":
                        RunWordParagraphNumberExamples();
                        break;

                    case "18":
                        RunWordParagraphRomanAlphabeticExamples();
                        break;

                    case "19":
                        RunWordParagraphFrameExamples();
                        break;

                    case "20":
                        RunWordListExamples();
                        break;

                    case "21":
                        RunWordTableExamples();
                        break;

                    case "22":
                        RunWordImageExamples();
                        break;

                    case "23":
                        RunWordShapeExamples();
                        break;

                    case "24":
                        RunWordGroupShapeConnectorExamples();
                        break;

                    case "25":
                        RunWordMetadataExamples();
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
            Console.WriteLine("14. Word Paragraph Examples");
            Console.WriteLine("15. Word Paragraph Alignment Examples");
            Console.WriteLine("16. Word Paragraph Indentation Examples");
            Console.WriteLine("17. Word Paragraph Number Examples");
            Console.WriteLine("18. Word Paragraph Roman Alphabetic Examples");
            Console.WriteLine("19. Word Paragraph Frame Examples");
            Console.WriteLine("20. Word List Examples");
            Console.WriteLine("21. Word Table Examples");
            Console.WriteLine("22. Word Image Examples");
            Console.WriteLine("23. Word Shape Examples");
            Console.WriteLine("24. Word Group Shape Connector Examples");
            Console.WriteLine("25. Word Metadata Examples");
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

        static void RunWordParagraphExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Paragraph Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ParagraphExamples();

            Console.WriteLine("\n1. CreateWordParagraphs...");
            Console.WriteLine("=============================");
            example.CreateWordParagraphs();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadWordParagraphs...");
            Console.WriteLine("=============================");
            example.ReadWordParagraphs();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyWordParagraphs...");
            Console.WriteLine("=============================");
            example.ModifyWordParagraphs();
            Console.WriteLine("=============================");
        }

        static void RunWordParagraphAlignmentExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Paragraph Alignment Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ParagraphAlignmentExamples();

            Console.WriteLine("\n1. CreateAlignment...");
            Console.WriteLine("=============================");
            example.CreateAlignment();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadAlignment...");
            Console.WriteLine("=============================");
            example.ReadAlignment();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyAlignment...");
            Console.WriteLine("=============================");
            example.ModifyAlignment();
            Console.WriteLine("=============================");
        }

        static void RunWordParagraphIndentExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Paragraph Indentation Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ParagraphIndentExamples();

            Console.WriteLine("\n1. CreateIndent...");
            Console.WriteLine("=============================");
            example.CreateIndent();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadIndent...");
            Console.WriteLine("=============================");
            example.ReadIndent();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyIndent...");
            Console.WriteLine("=============================");
            example.ModifyIndent();
            Console.WriteLine("=============================");
        }

        static void RunWordParagraphNumberExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Paragraph Number Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ParagraphNumberExamples();

            Console.WriteLine("\n1. CreateNumberedParagraphs...");
            Console.WriteLine("=============================");
            example.CreateNumberedParagraphs();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadNumberedParagraphs...");
            Console.WriteLine("=============================");
            example.ReadNumberedParagraphs();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyNumberedParagraphs...");
            Console.WriteLine("=============================");
            example.ModifyNumberedParagraphs();
            Console.WriteLine("=============================");
        }

        static void RunWordParagraphRomanAlphabeticExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Paragraph Roman Alphabetic Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ParagraphRomanAlphabeticExamples();

            Console.WriteLine("\n1. CreateRomanAlphabeticParagraphs...");
            Console.WriteLine("=============================");
            example.CreateRomanAlphabeticParagraphs();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadRomanAlphabeticParagraphs...");
            Console.WriteLine("=============================");
            example.ReadRomanAlphabeticParagraphs();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyRomanAlphabeticParagraphs...");
            Console.WriteLine("=============================");
            example.ModifyRomanAlphabeticParagraphs();
            Console.WriteLine("=============================");
        }

        static void RunWordParagraphFrameExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Paragraph Frame Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ParagraphFrameExamples();

            Console.WriteLine("\n1. CreateParagraphFrames...");
            Console.WriteLine("=============================");
            example.CreateParagraphsFrames();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadParagraphFrames...");
            Console.WriteLine("=============================");
            example.ReadParagraphsFrames();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyParagraphFrames...");
            Console.WriteLine("=============================");
            example.ModifyParagraphsFrames();
            Console.WriteLine("=============================");
        }

        static void RunWordListExamples()
        {
            Console.Clear();
            Console.WriteLine("Word List Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ListExamples();

            Console.WriteLine("\n1. CreateMultilevelLists...");
            Console.WriteLine("=============================");
            example.CreateMultilevelLists();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadMultilevelLists...");
            Console.WriteLine("=============================");
            example.ReadMultilevelLists();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyMultilevelLists...");
            Console.WriteLine("=============================");
            example.ModifyMultilevelLists();
            Console.WriteLine("=============================");
        }

        static void RunWordTableExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Table Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.TableExamples();

            Console.WriteLine("\n1. CreateWordDocumentWithTables...");
            Console.WriteLine("=============================");
            example.CreateWordDocumentWithTables();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadTablesInWordDocument...");
            Console.WriteLine("=============================");
            example.ReadTablesInWordDocument();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyTablesInWordDocument...");
            Console.WriteLine("=============================");
            example.ModifyTablesInWordDocument();
            Console.WriteLine("=============================");
        }

        static void RunWordImageExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Image Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ImageExamples();

            Console.WriteLine("\n1. CreateWordDocumentWithImages...");
            Console.WriteLine("=============================");
            example.CreateWordDocumentWithImages();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadImagesInWordDocument...");
            Console.WriteLine("=============================");
            example.ReadImagesInWordDocument();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyImagesInWordDocument...");
            Console.WriteLine("=============================");
            example.ModifyImagesInWordDocument();
            Console.WriteLine("=============================");
        }

        static void RunWordShapeExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Shape Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.ShapeExamples();

            Console.WriteLine("\n1. CreateShapes...");
            Console.WriteLine("=============================");
            example.CreateShapes();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadShapes...");
            Console.WriteLine("=============================");
            example.ReadShapes();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyShapes...");
            Console.WriteLine("=============================");
            example.ModifyShapes();
            Console.WriteLine("=============================");
        }

        static void RunWordGroupShapeConnectorExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Group Shape Connector Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.GroupShapeConnectorExamples();

            Console.WriteLine("\n1. CreateGroupShapes...");
            Console.WriteLine("=============================");
            example.CreateGroupShapes();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadGroupShapes...");
            Console.WriteLine("=============================");
            example.ReadGroupShapes();
            Console.WriteLine("=============================");
        }

        static void RunWordMetadataExamples()
        {
            Console.Clear();
            Console.WriteLine("Word Metadata Examples");
            Console.WriteLine("===============");

            var example = new Openize.Words.Examples.MetadataExamples();

            Console.WriteLine("\n1. CreateMetadata...");
            Console.WriteLine("=============================");
            example.CreateMetadata();
            Console.WriteLine("=============================");

            Console.WriteLine("\n2. ReadMetadata...");
            Console.WriteLine("=============================");
            example.ReadMetadata();
            Console.WriteLine("=============================");

            Console.WriteLine("\n3. ModifyMetadata...");
            Console.WriteLine("=============================");
            example.ModifyMetadata();
            Console.WriteLine("=============================");
        }
    }
}