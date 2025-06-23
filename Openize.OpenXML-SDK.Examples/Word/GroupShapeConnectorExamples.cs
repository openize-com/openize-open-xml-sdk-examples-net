namespace Openize.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word tables
    /// using the <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Word/Shape/Group at the root of your project.
    /// // Check reference for more options and details.
    /// var groupShapeExamples = new Openize.Words.Examples.GroupShapeConnectorExamples();
    /// // Creates a word document with shapes and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// groupShapeExamples.CreateGroupShapes();
    /// // Reads shapes from the specified Word Document and displays shape attributes.
    /// // Check reference for more options and details.
    /// groupShapeExamples.ReadGroupShapes();
    /// </code>
    /// </example>
    public class GroupShapeConnectorExamples
    {
        private const string docsDirectory = "../../../Documents/Word/Shape/Group";
        /// <summary>
        /// Initializes a new instance of the <see cref="GroupShapeConnectorExamples"/> class.
        /// Prepares the directory 'Documents/Word/Shape/Group' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public GroupShapeConnectorExamples()
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
        /// Generates 5(rows) x 3(cols) tables with table styles defined by the Word document template.
        /// Appends each table to the body of the word document.
        /// Saves the newly created word document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Table' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordGroupShapes.docx").
        /// </param>
        public void CreateGroupShapes(string documentDirectory = docsDirectory,
            string filename = "WordGroupShapes.docx")
        {
            try
            {

                // Initialize a new word document with the default template
                var doc = new Openize.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new Openize.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Instantiate shape element with diamond and coordinates/size.
                var diamond = new Openize.Words.IElements.Shape(0, 0, 200, 200,
                                IElements.ShapeType.Diamond);

                // Instantiate shape element with oval and coordinates/size.
                var oval = new Openize.Words.IElements.Shape(300, 0, 200, 200,
                                IElements.ShapeType.Ellipse);

                // Group diamond and oval shapes with an auto connector
                var groupShape = new Openize.Words.IElements.GroupShape(diamond, oval);

                // Add group shape to the word document.
                body.AppendChild(groupShape);
                System.Console.WriteLine("Group shape (diamond and oval with auto connector) added");

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
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a>.
        /// Traverses through shapes of the Word document.
        /// Reads and displays properties of the shape.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Word/Shape/Group' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordGroupShapes.docx").
        /// </param>
        public void ReadGroupShapes(string documentDirectory = docsDirectory,
            string filename = "WordGroupShapes.docx")
        {
            try
            {
                // Load the Word Document.
                var doc = new Openize.Words.Document($"{documentDirectory}/{filename}");
                // Initialize the body with the loaded document.
                var body = new Openize.Words.Body(doc);

                // Load all shapes with the document
                var groupShapes = body.GroupShapes;

                // Initialize the shape counter
                var groupShapeNumber = 0;

                // Traverse through each shape and display its properties
                foreach (var groupShape in groupShapes)
                {
                    groupShapeNumber++;
                    System.Console.WriteLine("Group Shape Number : " + groupShapeNumber);
                    System.Console.WriteLine("Shape 1 in Group Shape Number " + groupShapeNumber);
                    System.Console.WriteLine("------ Type:" + groupShape.Shape1.Type);
                    System.Console.WriteLine("------ Id : " + groupShape.Shape1.ElementId);
                    System.Console.WriteLine("------ X : " + groupShape.Shape1.X);
                    System.Console.WriteLine("------ Y : " + groupShape.Shape1.Y);
                    System.Console.WriteLine("------ Width : " + groupShape.Shape1.Width);
                    System.Console.WriteLine("------ Height : " + groupShape.Shape1.Height);

                    System.Console.WriteLine("Shape 2 in Group Shape Number " + groupShapeNumber);
                    System.Console.WriteLine("------ Type:" + groupShape.Shape2.Type);
                    System.Console.WriteLine("------ Id : " + groupShape.Shape2.ElementId);
                    System.Console.WriteLine("------ X : " + groupShape.Shape2.X);
                    System.Console.WriteLine("------ Y : " + groupShape.Shape2.Y);
                    System.Console.WriteLine("------ Width : " + groupShape.Shape2.Width);
                    System.Console.WriteLine("------ Height : " + groupShape.Shape2.Height);
                }
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
    }
}