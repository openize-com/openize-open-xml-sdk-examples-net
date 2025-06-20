using System;

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
                DisplayMainMenu();
                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        ExcelProgram.Run();
                        break;

                    case "2":
                        WordProgram.Run();
                        break;

                    case "3":
                        PowerPointProgram.Run();
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
                    Console.WriteLine("\nPress any key to return to the main menu...");
                    Console.ReadKey();
                    Console.Clear();
                }
            }

            Console.WriteLine("\nThank you for exploring the Openize.OpenXML-SDK Examples!");
        }

        static void DisplayMainMenu()
        {
            Console.WriteLine("Choose a product to explore:");
            Console.WriteLine("1. Excel Examples");
            Console.WriteLine("2. Word Examples");
            Console.WriteLine("3. PowerPoint Examples");
            Console.WriteLine("0. Exit");
            Console.Write("\nEnter your choice: ");
        }
    }
}