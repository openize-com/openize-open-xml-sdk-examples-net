namespace Openize.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word tables
    /// using the <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Word/Shape at the root of your project.
    /// // Check reference for more options and details.
    /// var shapeExamples = new Openize.Words.Examples.ShapeFillExamples();
    /// // Creates a word document having shapes with fill options and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// shapeExamples.CreateFillShapes();
    /// // Creates a word document having group shapes with fill options and saves word document to the specified
    /// // directory.Check reference for more options and details.
    /// shapeExamples.CreateFillGroupShapes();
    /// // Modifies shapes in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// shapeExamples.ModifyShapes();
    /// </code>
    /// </example>
    public class ShapeFillExamples
    {
        private const string docsDirectory = "../../../Documents/Word/Shape";
        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeExamples"/> class.
        /// Prepares the directory 'Documents/Word/Shape' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ShapeFillExamples()
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
        /// Generates hexagone shapes with pattern fill.
        /// Appends shape to the body of the word document.
        /// Saves the newly created word document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Word/Shape' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordShapesFill.docx").
        /// </param>
        public void CreateFillShapes(string documentDirectory = docsDirectory,
            string filename = "WordShapesFill.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new Openize.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new Openize.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Define two colors for filltype
                var shapeColors = new Openize.Words.IElements.ShapeFillColors();
                shapeColors.Color1 = Openize.Words.IElements.Colors.Red;
                shapeColors.Color2 = Openize.Words.IElements.Colors.Purple;

                // Create Shape with Pattern Fill with colors defined above
                var shape = new Openize.Words.IElements.Shape(100, 100, 400, 400,
                                Openize.Words.IElements.ShapeType.Hexagone,
                                Openize.Words.IElements.ShapeFillType.Pattern,
                                shapeColors);

                // Add shape to the word document.
                body.AppendChild(shape);
                System.Console.WriteLine("Hexagone shape added with Pattern Fill");

                // Save the newly created Word Document.
                doc.Save($"{documentDirectory}/{filename}");
                System.Console.WriteLine($"Word Document {filename} Created. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Creates a new Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a>.
        /// Generates diamond shapes with gradient fill.
        /// Generates ellipse shape with pattern fill
        /// Groups diamond and ellipse shape with auto connector
        /// Appends group shape to the body of the word document.
        /// Saves the newly created word document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Word/Shape' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordGroupShapesFill.docx").
        /// </param>
        public void CreateFillGroupShapes(string documentDirectory = docsDirectory,
            string filename = "WordGroupShapesFill.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new Openize.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new Openize.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Create diamond shape with Gradient fill
                var diamond = new Openize.Words.IElements.Shape(0, 0, 200, 200,
                          Openize.Words.IElements.ShapeType.Diamond,
                          Openize.Words.IElements.ShapeFillType.Gradient,
                          new Openize.Words.IElements.ShapeFillColors());

                // Create ellipse shape with Pattern fill
                var oval = new Openize.Words.IElements.Shape(300, 0, 200, 200,
                                          Openize.Words.IElements.ShapeType.Ellipse,
                                          Openize.Words.IElements.ShapeFillType.Pattern,
                                          new Openize.Words.IElements.ShapeFillColors());

                // Group diamond and ellipse shapes with auto connector
                var groupShape = new Openize.Words.IElements.GroupShape(diamond, oval);

                // Add shape to the word document.
                body.AppendChild(groupShape);
                System.Console.WriteLine("Group shape added consisting of diamond and ellipse with fill options");

                // Save the newly created Word Document.
                doc.Save($"{documentDirectory}/{filename}");
                System.Console.WriteLine($"Word Document {filename} Created. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
    }
}
