using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Archsheerary
{
    public partial class OOXML
    {
        /// <summary>
        /// Extract data from an OOXML spreadsheet to a file in directory
        /// </summary>
        public class Extract
        {
            /// <summary>
            /// Extract embedded objects to a directory
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <returns>List of extracted embedded objects</returns>
            public List<DataTypes.EmbeddedObjects> EmbeddedObjects(string input_filepath, string output_folder)
            {
                // Create new directory if it does not exist
                Directory.CreateDirectory(output_folder);

                // Create new list
                List<DataTypes.EmbeddedObjects> results = new List<DataTypes.EmbeddedObjects>();

                // Define data types
                List<EmbeddedObjectPart> ole = new List<EmbeddedObjectPart>();
                List<EmbeddedPackagePart> packages = new List<EmbeddedPackagePart>();
                List<ImagePart> emf = new List<ImagePart>();
                List<ImagePart> images = new List<ImagePart>();
                List<Model3DReferenceRelationshipPart> threeD = new List<Model3DReferenceRelationshipPart>();

                // Open the spreadsheet
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
                {
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        // Perform check
                        ole = worksheetPart.EmbeddedObjectParts.Distinct().ToList();
                        packages = worksheetPart.EmbeddedPackageParts.Distinct().ToList();
                        emf = worksheetPart.ImageParts.Distinct().ToList();
                        if (worksheetPart.DrawingsPart != null) // DrawingsPart needs a null check
                        {
                            images = worksheetPart.DrawingsPart.ImageParts.Distinct().ToList();
                        }
                        threeD = worksheetPart.Model3DReferenceRelationshipParts.Distinct().ToList();

                        // Extract each part
                        foreach (EmbeddedObjectPart part in ole)
                        {
                            // Create filename
                            int integer = 0;
                            string filename = part.Uri.ToString().Split("/").Last();
                            string output_filepath = output_folder + "\\" + filename;
                            while (File.Exists(output_filepath))
                            {
                                integer++;
                                output_filepath = output_folder + "\\" + filename +"_" + integer;
                            }

                            // Create extracted file
                            using (FileStream fileStream = File.Create(output_filepath))
                            {
                                Stream stream = part.GetStream();
                                stream.Seek(0, SeekOrigin.Begin);
                                stream.CopyTo(fileStream);
                            }

                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                        }
                        foreach (EmbeddedPackagePart part in packages)
                        {
                            // Create filename
                            int integer = 0;
                            string filename = part.Uri.ToString().Split("/").Last();
                            string output_filepath = output_folder + "\\" + filename;
                            while (File.Exists(output_filepath))
                            {
                                integer++;
                                output_filepath = output_folder + "\\" + filename + "_" + integer;
                            }

                            // Create extracted file
                            using (FileStream fileStream = File.Create(output_filepath))
                            {
                                Stream stream = part.GetStream();
                                stream.Seek(0, SeekOrigin.Begin);
                                stream.CopyTo(fileStream);
                            }

                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                        }
                        foreach (Model3DReferenceRelationshipPart part in threeD)
                        {
                            // Create filename
                            int integer = 0;
                            string filename = part.Uri.ToString().Split("/").Last();
                            string output_filepath = output_folder + "\\" + filename;
                            while (File.Exists(output_filepath))
                            {
                                integer++;
                                output_filepath = output_folder + "\\" + filename + "_" + integer;
                            }

                            // Create extracted file
                            using (FileStream fileStream = File.Create(output_filepath))
                            {
                                Stream stream = part.GetStream();
                                stream.Seek(0, SeekOrigin.Begin);
                                stream.CopyTo(fileStream);
                            }

                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                        }
                        foreach (ImagePart part in emf)
                        {
                            // Create filename
                            int integer = 0;
                            string filename = part.Uri.ToString().Split("/").Last();
                            string output_filepath = output_folder + "\\" + filename;
                            while (File.Exists(output_filepath))
                            {
                                integer++;
                                output_filepath = output_folder + "\\" + filename + "_" + integer;
                            }

                            // Create extracted file
                            using (FileStream fileStream = File.Create(output_filepath))
                            {
                                Stream stream = part.GetStream();
                                stream.Seek(0, SeekOrigin.Begin);
                                stream.CopyTo(fileStream);
                            }

                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                        }
                        foreach (ImagePart part in images)
                        {
                            // Create filename
                            int integer = 0;
                            string filename = part.Uri.ToString().Split("/").Last();
                            string output_filepath = output_folder + "\\" + filename;
                            while (File.Exists(output_filepath))
                            {
                                integer++;
                                output_filepath = output_folder + "\\" + filename + "_" + integer;
                            }

                            // Create extracted file
                            using (FileStream fileStream = File.Create(output_filepath))
                            {
                                Stream stream = part.GetStream();
                                stream.Seek(0, SeekOrigin.Begin);
                                stream.CopyTo(fileStream);
                            }

                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                        }
                    }
                }
                // Delete new folder if no objects were copied to it
                if (Directory.GetFiles(output_folder).Length == 0)
                {
                    Directory.Delete(output_folder);
                }
                return results;
            }

            /// <summary>
            /// Extract external object references.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <returns>List of extracted external object references</returns>
            /// <exception cref="IOException">Thrown if file to be extracted cannot be found.</exception>
            public static List<DataTypes.ExternalObjects> ExternalObjects(string filepath, string output_folder)
            {
                // Create new directory if it does not exist
                Directory.CreateDirectory(output_folder);

                List<DataTypes.ExternalObjects> results = new List<DataTypes.ExternalObjects>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                    foreach (ExternalWorkbookPart extWbPart in extWbParts)
                    {
                        List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList();
                        foreach (ExternalRelationship extrel in extrels)
                        {
                            string output_filepath = output_folder + "\\" + extrel.Uri.ToString().Split("/").Last();
                            try
                            {
                                // Copy external file to subfolder
                                File.Copy(extrel.Uri.ToString(), output_filepath);

                                // Add to list
                                results.Add(new DataTypes.ExternalObjects() { Target = extrel.Uri.ToString(), RelationshipType = extrel.RelationshipType, IsExternal = extrel.IsExternal, Container = extrel.Container.ToString(), ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                            }
                            catch (System.IO.IOException)
                            {
                                // Add as failed to list
                                results.Add(new DataTypes.ExternalObjects() { Target = extrel.Uri.ToString(), RelationshipType = extrel.RelationshipType, IsExternal = extrel.IsExternal, Container = extrel.Container.ToString(), ExtractedFilepath = output_filepath, Action = DataTypes.ActionFailed });
                            }
                        }
                    }
                }
                // Delete new folder, if no objects were copied to it
                if (Directory.GetFiles(output_folder).Length == 0)
                {
                    Directory.Delete(output_folder);
                }
                return results;
            }

            /// <summary>
            /// Extract file property information to a text file
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <returns>List of extracted file property information and saved in file Metadata.txt</returns>
            public List<DataTypes.FilePropertyInformation> FilePropertyInformation(string input_filepath, string output_folder)
            {
                bool found = false;

                // Create new list
                List<DataTypes.FilePropertyInformation> results = new List<DataTypes.FilePropertyInformation>();

                // Create new folder if it does not exist
                Directory.CreateDirectory(output_folder);

                // Set output filepath
                string output_filepath = $"{output_folder}\\Metadata.txt";

                // Open the spreadsheet
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
                {
                    PackageProperties property = spreadsheet.PackageProperties;

                    // Create metadata file
                    using (StreamWriter w = File.AppendText(output_filepath))
                    {
                        w.WriteLine("FILE PROPERTIES INFORMATION");
                        w.WriteLine($"INFORMATION FROM: {input_filepath}");
                        w.WriteLine("---");

                        // Write information to metadata file
                        if (property.Creator != null)
                        {
                            w.WriteLine($"CREATOR: {property.Creator}");
                            found = true;
                        }
                        if (property.Title != null)
                        {
                            w.WriteLine($"TITLE: {property.Title}");
                            found = true;
                        }
                        if (property.Subject != null)
                        {
                            w.WriteLine($"SUBJECT: {property.Subject}");
                            found = true;
                        }
                        if (property.Description != null)
                        {
                            w.WriteLine($"DESCRIPTION: {property.Description}");
                            found = true;
                        }
                        if (property.Keywords != null)
                        {
                            w.WriteLine($"KEYWORDS: {property.Keywords}");
                            found = true;
                        }
                        if (property.Category != null)
                        {
                            w.WriteLine($"CATEGORY: {property.Category}");
                            found = true;
                        }
                        if (property.LastModifiedBy != null)
                        {
                            w.WriteLine($"LAST MODIFIED BY: {property.LastModifiedBy}");
                            found = true;
                        }
                        // Add to list
                        results.Add(new DataTypes.FilePropertyInformation() { Author = property.Creator, Title = property.Title, Subject = property.Subject, Description = property.Description, Keywords = property.Keywords, Category = property.Category, LastModifiedBy = property.LastModifiedBy, FilePropertyInfoFound = found, ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                    }
                }
                // Delete new folder, if no objects were copied to it
                if (Directory.GetFiles(output_folder).Length == 0)
                {
                    Directory.Delete(output_folder);
                }
                return results;
            }

            /// <summary>
            /// Extract all cell hyperlinks to an external file
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <returns>List of extracted hyperlinks and saved in file Metadata.txt</returns>
            public List<DataTypes.Hyperlinks> Extract_Hyperlinks(string input_filepath, string output_folder)
            {
                // Create new list
                List<DataTypes.Hyperlinks> results = new List<DataTypes.Hyperlinks>();

                // Create new folder if it does not exist
                Directory.CreateDirectory(output_folder);

                // Set output filepath
                string output_filepath = $"{output_folder}\\Metadata.txt";

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
                {
                    List<HyperlinkRelationship> hyperlinks = spreadsheet
                        .GetAllParts()
                        .SelectMany(p => p.HyperlinkRelationships)
                        .ToList();

                    // Create metadata file
                    using (StreamWriter w = File.AppendText(output_filepath))
                    {
                        w.WriteLine("---");
                        w.WriteLine("EXTRACTED HYPERLINKS");
                        w.WriteLine("---");

                        foreach (HyperlinkRelationship hyperlink in hyperlinks)
                        {
                            // Write information to metadata file
                            w.WriteLine(hyperlink.Uri);
                            // Add to list
                            results.Add(new DataTypes.Hyperlinks() { URL = hyperlink.Uri.ToString(), ExtractedFilepath = output_filepath, Action = DataTypes.ActionExtracted });
                        }
                    }
                }
                return results;
            }
        }
    }
}
