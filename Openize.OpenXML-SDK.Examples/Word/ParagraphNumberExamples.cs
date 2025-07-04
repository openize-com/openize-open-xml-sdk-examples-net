﻿using System;
namespace Openize.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying numbered paragraphs in DOCX word
    /// using the <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Paragraph/Numbering at the root of your project.
    /// // Check reference for more options and details.
    /// var paragraphNumberExamples = new ParagraphNumberExamples();
    /// // Creates a word document with paragraphs having various numbering levels and saves word
    /// // document to the specified directory. Check reference for more options and details.
    /// paragraphNumberExamples.CreateNumberedParagraphs();
    /// // Reads Paragraphs from the specified Word Document and displays plain text alongwith numbering info.
    /// // Check reference for more options and details.
    /// paragraphNumberExamples.ReadNumberedParagraphs();
    /// // Modifies Paragraph's numbering in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// paragraphNumberExamples.ModifyNumberedParagraphs();
    /// </code>
    /// </example>
    public class ParagraphNumberExamples
    {
        private const string docsDirectory = "../../../Documents/Word/Paragraph/Numbering";
        /// <summary>
        /// Initializes a new instance of the <see cref="ParagraphNumberExamples"/> class.
        /// Prepares the directory 'Documents/Paragraph/Numbering' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ParagraphNumberExamples()
        {
            if (!System.IO.Directory.Exists(docsDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(docsDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' " +
                    $"created successfully.");
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(docsDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' " +
                    $"cleaned up.");
            }
        }
        /// <summary>
        /// Creates a new Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a>.
        /// Generates numbered paragraphs with nested levels.
        /// Saves the newly created Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Paragraph/Numbering' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordParagraphsNumbered.docx").
        /// </param>
        public void CreateNumberedParagraphs(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsNumbered.docx")
        {
            try
            {
                // Initialize docx document
                Openize.Words.Document doc = new Openize.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize document body
                var body = new Openize.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Initialize a paragraph
                var para = new Openize.Words.IElements.Paragraph();

                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "This document is generated by Openize.OpenXML-SDK." });
                // Append paragraph to the document body
                body.AppendChild(para);

                // Reset paragraph
                para = new Openize.Words.IElements.Paragraph();
                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "Below are numbered paragraphs:" });
                // Append paragraph to the document body
                body.AppendChild(para);

                // Reset paragraph
                para = new Openize.Words.IElements.Paragraph { Style = "ListParagraph" };
                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "First numbered  at first level" });
                // Set numbering for the paragraph
                para.NumberingId = 1;
                para.IsNumbered = true;
                para.NumberingLevel = 1;

                // Append paragraph to the document body
                body.AppendChild(para);

                // Reset paragraph
                para = new Openize.Words.IElements.Paragraph { Style = "ListParagraph" };
                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "First numbered at second level" });
                // Set numbering for the paragraph
                para.NumberingId = 1;
                para.IsNumbered = true;
                para.NumberingLevel = 2;
                // Append paragraph to the document body
                body.AppendChild(para);

                // Reset paragraph
                para = new Openize.Words.IElements.Paragraph { Style = "ListParagraph" };
                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "Second numbered at second level" });
                // Set numbering for the paragraph
                para.NumberingId = 1;
                para.IsNumbered = true;
                para.NumberingLevel = 2;
                // Append paragraph to the document body
                body.AppendChild(para);

                // Reset paragraph
                para = new Openize.Words.IElements.Paragraph { Style = "ListParagraph" };
                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "Second numbered at first level" });
                // Set numbering for the paragraph
                para.NumberingId = 1;
                para.IsNumbered = true;
                para.NumberingLevel = 1;
                // Append paragraph to the document body
                body.AppendChild(para);

                // Reset paragraph
                para = new Openize.Words.IElements.Paragraph();
                // Add a run to paragraph
                para.AddRun(new Openize.Words.IElements.Run
                { Text = "The document ends here..." });
                // Append paragraph to the doucment body
                body.AppendChild(para);

                // Save docx document to the disk
                doc.Save($"{documentDirectory}/{filename}");
                System.Console.WriteLine($"Word Document {filename} Created. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");

                // The resulting docx document should be like this: https://imgur.com/vjTJxw5
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a>.
        /// Traverses paragraphs and displays its text, numbering and level.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph/Numbering' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordParagraphsNumbered.docx").
        /// </param>
        public void ReadNumberedParagraphs(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsNumbered.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new Openize.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new Openize.Words.Body(doc);

                //System.Collections.Generic.List<Openize.Words.IElements.Paragraph>
                //  paragraphs = body.Paragraphs;

                var paragraphs = body.Paragraphs;

                foreach (Openize.Words.IElements.Paragraph paragraph in paragraphs)
                {
                    System.Console.WriteLine($"Paragraph Text : {paragraph.Text}");
                    System.Console.WriteLine($"Paragraph NumberingId : {paragraph.NumberingId}");
                    System.Console.WriteLine($"Paragraph Numbered? : {paragraph.IsNumbered}");
                    System.Console.WriteLine($"Paragraph Roman? : {paragraph.IsRoman}");
                    System.Console.WriteLine($"Paragraph AlphabeticNumber? : {paragraph.IsAlphabeticNumber}");
                    System.Console.WriteLine($"Paragraph Numbering Level : {paragraph.NumberingLevel}");
                }
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a>.
        /// Traverses through all paragraphs within the document.
        /// If numbered, modifies paragraphs by appending ' (numering removed)' with italic format
        /// and paragraph number is removed.
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph/Numbering' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordParagraphsNumbered.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordParagraphsNumbered.docx").
        /// </param>
        public void ModifyNumberedParagraphs(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsNumbered.docx",
            string filenameModified = "ModifiedWordParagraphsNumbered.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new Openize.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new Openize.Words.Body(doc);

                //foreach (Openize.Words.IElements.Paragraph paragraph in body.Paragraphs)
                foreach (var paragraph in body.Paragraphs)
                {
                    if (paragraph.Style == "ListParagraph")
                    {
                        paragraph.Style = "Normal";
                        paragraph.AddRun(new Openize.Words.IElements.Run
                        { Text = " (numbering removed)", Italic = true });
                        doc.Update(paragraph);
                    }
                }

                // Save the modified Word Document
                doc.Save($"{documentDirectory}/{filenameModified}");
                System.Console.WriteLine($"Word Document {filename} Modified and Saved As " +
                    $"{filenameModified}. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
    }
}
