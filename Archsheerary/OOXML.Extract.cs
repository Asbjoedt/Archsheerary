using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
            public void EmbeddedObjects(string input_filepath, string output_folder)
            {
                // Create new directory if it does not exist
                Directory.CreateDirectory(output_folder);

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
                        }
                    }
                }
            }

            /// <summary>
            /// Extract file property information to a text file
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            public void FilePropertyInformation(string input_filepath, string output_folder)
            {
                // Create new folder if it does not exist
                Directory.CreateDirectory(output_folder);

                // Create filename if it does not exist
                int integer = 1;
                string output_filepath = $"{output_folder}\\FilePropertyInformation{integer}.txt";
                while (File.Exists(output_filepath))
                {
                    integer++;
                    output_filepath = $"{output_folder}\\FilePropertyInformation{integer}.txt";
                }

                // Open the spreadsheet
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
                {
                    PackageProperties property = spreadsheet.Package.PackageProperties;

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
                        }
                        if (property.Title != null)
                        {
                            w.WriteLine($"TITLE: {property.Title}");
                        }
                        if (property.Subject != null)
                        {
                            w.WriteLine($"SUBJECT: {property.Subject}");
                        }
                        if (property.Description != null)
                        {
                            w.WriteLine($"DESCRIPTION: {property.Description}");
                        }
                        if (property.Keywords != null)
                        {
                            w.WriteLine($"KEYWORDS: {property.Keywords}");
                        }
                        if (property.Category != null)
                        {
                            w.WriteLine($"CATEGORY: {property.Category}");
                        }
                        if (property.LastModifiedBy != null)
                        {
                            w.WriteLine($"LAST MODIFIED BY: {property.LastModifiedBy}");
                        }
                    }
                }
            }
        }
    }
}
