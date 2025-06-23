namespace Openize.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word tables
    /// using the <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Word/Metadata at the root of your project.
    /// // Check reference for more options and details.
    /// var shapeExamples = new Openize.Words.Examples.ShapeExamples();
    /// // Creates a word document with shapes and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// shapeExamples.CreateShapes();
    /// // Reads shapes from the specified Word Document and displays shape attributes.
    /// // Check reference for more options and details.
    /// shapeExamples.ReadShapes();
    /// // Modifies shapes in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// shapeExamples.ModifyShapes();
    /// </code>
    /// </example>
    public class MetadataExamples
    {
        private const string docsDirectory = "../../../Documents/Word/Metadata";
        /// <summary>
        /// Initializes a new instance of the <see cref="MetadataExamples"/> class.
        /// Prepares the directory 'Documents/Word/Metadata' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public MetadataExamples()
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
        /// Generates several metadata values including title, subject, description and so on.
        /// Sets metadata of the word document.
        /// Saves the newly created word document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the
        /// 'Documents/Word/Metadata' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordMetadata.docx").
        /// </param>
        public void CreateMetadata(string documentDirectory = docsDirectory,
            string filename = "WordMetadata.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new Openize.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new Openize.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Instantiate document metadata
                var docMetadata = new Openize.Words.DocumentProperties();

                // Define metadata attributes
                docMetadata.Title = "My Title";
                docMetadata.Subject = "My Subject";
                docMetadata.Description = "My Description";
                docMetadata.Keywords = "Openize.OpenXML-SDK";
                docMetadata.Creator = "Openize.OpenXML-SDK for .NET";
                docMetadata.LastModifiedBy = "Openize.OpenXML-SDK for .NET";
                docMetadata.Revision = "1";
                var currentTime = System.DateTime.UtcNow;
                docMetadata.Created = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
                docMetadata.Modified = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");

                // Set Metadata
                doc.SetDocumentProperties(docMetadata);

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
        /// Geta metadata properties of the Word document.
        /// Reads and displays metadata values of the document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Word/Metadata' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordMetadata.docx").
        /// </param>
        public void ReadMetadata(string documentDirectory = docsDirectory,
            string filename = "WordMetadata.docx")
        {
            try
            {
                // Load the Word Document.
                var doc = new Openize.Words.Document($"{documentDirectory}/{filename}");
                // Initialize the body with the loaded document.
                var body = new Openize.Words.Body(doc);

                // Get document metadata
                var coreprops = doc.GetDocumentProperties();

                // Display document metadata
                System.Console.WriteLine("Creator: " + coreprops.Creator);
                System.Console.WriteLine("Keywords: " + coreprops.Keywords);
                System.Console.WriteLine("Title: " + coreprops.Title);
                System.Console.WriteLine("Subject: " + coreprops.Subject);
                System.Console.WriteLine("Description: " + coreprops.Description);
                System.Console.WriteLine("LastModifiedBy: " + coreprops.LastModifiedBy);
                System.Console.WriteLine("Revision: " + coreprops.Revision);
                System.Console.WriteLine("Created: " + coreprops.Created);
                System.Console.WriteLine("Modified: " + coreprops.Modified);
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/Openize.OpenXML-SDK">Openize.OpenXML-SDK</a>.
        /// Gets the metadata properties of the Word document.
        /// Updates metadata values
        /// Sets the new metadata values.
        /// Displays old and new values.
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the
        /// 'Documents/Word/Metadata' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordShapes.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordShapes.docx").
        /// </param>
        public void ModifyMetadata(string documentDirectory = docsDirectory,
            string filename = "WordMetadata.docx", string filenameModified = "ModifiedWordMetadata.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new Openize.Words.Document($"{documentDirectory}/{filename}");
                // Initialize the body with the loaded document. 
                var body = new Openize.Words.Body(doc);

                // Get current metadata
                var oldProps = doc.GetDocumentProperties();
                var props = new Openize.Words.DocumentProperties();

                // Helper to print and update
                void UpdateProperty<T>(string name, T oldValue, T newValue, System.Action<T> setValue)
                {
                    System.Console.WriteLine($"Old {name} : {oldValue}");
                    setValue(newValue);
                    System.Console.WriteLine($"New {name} : {newValue}");
                }

                // Update metadata
                UpdateProperty("Title", oldProps.Title,
                    "Updated Title", val => props.Title = val);
                UpdateProperty("Subject", oldProps.Subject,
                    "Updated Subject", val => props.Subject = val);
                UpdateProperty("Description", oldProps.Description,
                    "Updated Description", val => props.Description = val);
                UpdateProperty("Creator", oldProps.Creator,
                    "Updated Creator", val => props.Creator = val);
                UpdateProperty("Keywords", oldProps.Keywords,
                    "Updated.Keyword", val => props.Keywords = val);
                UpdateProperty("LastModifiedBy", oldProps.LastModifiedBy,
                    "Updated.Openize.OpenXML-SDK", val => props.LastModifiedBy = val);
                UpdateProperty("Revision", oldProps.Revision,
                    "Version.Revised", val => props.Revision = val);
                UpdateProperty("Created", oldProps.Created,
                    oldProps.Created, val => props.Created = val);

                var currentTime = System.DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
                UpdateProperty("Modified", oldProps.Modified,
                    currentTime, val => props.Modified = val);

                // Set updated metadata
                doc.SetDocumentProperties(props);
                
                // Save the modified Word Document
                doc.Save($"{documentDirectory}/{filenameModified}");
                System.Console.WriteLine($"Word Document {filename} Modified and " +
                    $"Saved As {filenameModified}. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new Openize.Words.OpenizeException("An error occurred.", ex);
            }
        }
    }
}
