using Openize.OpenXML_SDK.Examples.PowerPoint;
using Openize.Slides.Examples;
using System;


namespace Openize.OpenXML_SDK.Examples.Usage
{
    public static class PowerPointProgram
    {
        public static void Run()
        {
            bool back = false;

            while (!back)
            {
                Console.Clear();
                DisplayPowerPointMenu();
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        RunSlideExamples();
                        break;
                    case "2":
                        RunSlideTextExamples();
                        break;
                    case "3":
                        RunSlideImageExamples();
                        break;
                    case "4":
                        RunSlideStyledListExamples();
                        break;
                    case "5":
                        RunSlideTableExamples();
                        break;
                    case "6":
                        RunSlideCommentExamples();
                        break;
                    case "7":
                        RunSlideCommentAuthorExamples();
                        break;
                    case "8":
                        RunSlideNotesExamples();
                        break;
                    case "9":
                        RunSlideRectangleExamples();
                        break;
                    case "10":
                        RunSlideCircleExamples();
                        break;
                    case "11":
                        RunSlideAnimationExamples();
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
                    Console.WriteLine("\nPress any key to return to the PowerPoint menu...");
                    Console.ReadKey();
                }
            }
        }

        private static void DisplayPowerPointMenu()
        {
            Console.WriteLine("PowerPoint Examples");
            Console.WriteLine("===================");
            Console.WriteLine("1. Slide Examples");
            Console.WriteLine("2. Slide Text Examples");
            Console.WriteLine("3. Slide Image Examples");
            Console.WriteLine("4. Slide Styled List Examples");
            Console.WriteLine("5. Slide Table Examples");
            Console.WriteLine("6. Slide Comment Examples");
            Console.WriteLine("7. Slide Comment Author Examples");
            Console.WriteLine("8. Slide Notes Examples");
            Console.WriteLine("9. Slide Rectangle Examples");
            Console.WriteLine("10. Slide Circle Examples");
            Console.WriteLine("11. Slide Animation Examples");
            Console.WriteLine("0. Back to Main Menu");
            Console.Write("\nEnter your choice: ");
        }

        // Add your PowerPoint example methods here
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

        static void RunSlideAnimationExamples()
        {
            Console.Clear();
            Console.WriteLine("Animation Examples");
            Console.WriteLine("==================");

            var example = new AnimationExamples();

            Console.WriteLine("\n1. Apply Zoom Animation...");
            example.ApplyZoomAnimation();

            Console.WriteLine("\n2. Apply FlyIn Animation...");
            example.ApplyFlyInAnimation();

            Console.WriteLine("\n3. Apply Spin Animation...");
            example.ApplySpinAnimation();

            Console.WriteLine("\n4. Apply FloatIn Animation...");
            example.ApplyFloatInAnimation();

            Console.WriteLine("\n5. Apply Bounce Animation...");
            example.ApplyBounceAnimation();
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