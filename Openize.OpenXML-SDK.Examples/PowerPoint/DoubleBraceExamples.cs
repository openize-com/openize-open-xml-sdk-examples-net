﻿using Openize.Slides;
using Openize.Slides.Common;
using System;
using System.Collections.Generic;



namespace Openize.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying DoubleBrace segments or shapes in a Presentation
    /// using the <a href="https://www.nuget.org/packages/Openize.Slides">Openize.Slides</a> library.
    /// </summary>
    public class DoubleBraceExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";

        /// <summary>
        /// Initializes a new instance of the <see cref="DoubleBraceExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public DoubleBraceExamples()
        {
            if (!System.IO.Directory.Exists(newDocsDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(newDocsDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(newDocsDirectory)}' " +
                    $"created successfully.");
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(newDocsDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(newDocsDirectory)}' " +
                    $"cleaned up.");
            }
        }
        /// <summary>
        /// This method adds DoubleBrace segment or shape in the silde of a new PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void DrawNewDoubleBraceShapeInNewSlide(string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Create an instance of DoubleBrace
                
                DoubleBrace DoubleBrace = new DoubleBrace();
                // Set height and width
                DoubleBrace.Width = 400.0;
                DoubleBrace.Height = 400.0;
                // Set Y position
                DoubleBrace.Y = 100.0;
                // First slide
                Slide slide = presentation.GetSlides()[1];
                // Add DoubleBrace shapes.
                slide.DrawDoubleBrace(DoubleBrace);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new Openize.Slides.Common.OpenizeException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method adds DoubleBrace segment or shape in the silde of a new PowerPoint presentation with animation.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void DrawNewDoubleBraceShapeWithAnimation(string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Create an instance of DoubleBrace

                DoubleBrace DoubleBrace = new DoubleBrace();
                // Set height and width
                DoubleBrace.Width = 400.0;
                DoubleBrace.Height = 400.0;
                // Set Y position
                DoubleBrace.Y = 100.0;
                DoubleBrace.Animation = Common.Enumerations.AnimationType.FlyIn;
                // First slide
                Slide slide = presentation.GetSlides()[1];
                // Add DoubleBrace shapes.
                slide.DrawDoubleBrace(DoubleBrace);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new Openize.Slides.Common.OpenizeException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method Sets the background color of a DoubleBrace shape
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void SetBackgroundColorOfDoubleBrace(string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get the slides
                Slide slide = presentation.GetSlides()[1];
                // Get 1st DoubleBrace
                DoubleBrace DoubleBrace = slide.DoubleBraces[0];
                // Set background of the DoubleBrace
                DoubleBrace.BackgroundColor = "289876";

                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new Openize.Slides.Common.OpenizeException("An error occurred.", ex);
            }
        }

        /// <summary>
        /// Remove DoubleBrace shape from an existing slide
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void RemoveDoubleBraceShapeExistingSlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
            // Get the slides
            Slide slide = presentation.GetSlides()[1];
            // Get 1st DoubleBrace
            DoubleBrace DoubleBrace = slide.DoubleBraces[0];
            // Remove DoubleBrace
            DoubleBrace.Remove();
            // Save the PPT or PPTX
            presentation.Save();

        }
    }
}
