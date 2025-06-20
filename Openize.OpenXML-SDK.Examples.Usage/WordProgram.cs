using Openize.OpenXML_SDK.Examples.Excel;
using Openize.Slides.Examples;
using System;

namespace Openize.OpenXML_SDK.Examples.Usage
{
    public static class WordProgram
    {
        public static void Run()
        {
            bool back = false;

            while (!back)
            {
                Console.Clear();
                DisplayWordMenu();
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        RunWordParagraphExamples();
                        break;

                    case "2":
                        RunWordParagraphAlignmentExamples();
                        break;

                    case "3":
                        RunWordParagraphIndentExamples();
                        break;

                    case "4":
                        RunWordParagraphNumberExamples();
                        break;

                    case "5":
                        RunWordParagraphRomanAlphabeticExamples();
                        break;

                    case "6":
                        RunWordParagraphFrameExamples();
                        break;

                    case "7":
                        RunWordListExamples();
                        break;

                    case "8":
                        RunWordTableExamples();
                        break;

                    case "9":
                        RunWordImageExamples();
                        break;

                    case "10":
                        RunWordShapeExamples();
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
                    Console.WriteLine("\nPress any key to return to the Word menu...");
                    Console.ReadKey();
                }
            }
        }

        private static void DisplayWordMenu()
        {
            Console.WriteLine("Word Examples");
            Console.WriteLine("=============");
            Console.WriteLine("Choose an example to run:");
            Console.WriteLine("1. Word Paragraph Examples");
            Console.WriteLine("2. Word Paragraph Alignment Examples");
            Console.WriteLine("3. Word Paragraph Indentation Examples");
            Console.WriteLine("4. Word Paragraph Number Examples");
            Console.WriteLine("5. Word Paragraph Roman Alphabetic Examples");
            Console.WriteLine("6. Word Paragraph Frame Examples");
            Console.WriteLine("7. Word List Examples");
            Console.WriteLine("8. Word Table Examples");
            Console.WriteLine("9. Word Image Examples");
            Console.WriteLine("10. Word Shape Examples");
            Console.WriteLine("0. Back to Main Menu");
            Console.Write("\nEnter your choice: ");
        }

        // Add your Word example methods here

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
    }
}