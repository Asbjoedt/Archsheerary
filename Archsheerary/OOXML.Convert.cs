using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Drawing;
using ImageMagick;

namespace Archsheerary
{
    public partial class OOXML
    {
        /// <summary>
        /// Collection of methods for converting Office Open XML spreadsheets
        /// </summary>
        public class Convert
        {
            /// <summary>
            /// Convert spreadsheet to XLSX Transitional conformance
            /// </summary>
            public static bool ToXLSXTransitional(string input_filepath, string output_filepath, bool set_normal_fileattributes)
            {
                bool convert_success = false;

                // If protected in file properties
                if (set_normal_fileattributes)
                {
                    File.SetAttributes(input_filepath, FileAttributes.Normal); // Remove file attributes on spreadsheet
                }

                // Convert spreadsheet
                byte[] byteArray = File.ReadAllBytes(input_filepath);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                    {
                        spreadsheet.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                    }
                    File.WriteAllBytes(output_filepath, stream.ToArray());
                }

                // Repair spreadsheet
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                convert_success = true;
                return convert_success;
            }
        }

        /// <summary>
        /// Convert embedded images to TIFF file format
        /// </summary>
        public void EmbeddedImagesToTiff(string filepath)
        {
            // Define data types
            List<ImagePart> emf = new List<ImagePart>();
            List<ImagePart> images = new List<ImagePart>();

            // Open the spreadsheet
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
            {
                IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                foreach (WorksheetPart worksheetPart in worksheetParts)
                {
                    // Perform check
                    emf = worksheetPart.ImageParts.Distinct().ToList();
                    if (worksheetPart.DrawingsPart != null) // DrawingsPart needs a null check
                    {
                        images = worksheetPart.DrawingsPart.ImageParts.Distinct().ToList();
                    }

                    // Convert each image
                    foreach (ImagePart part in emf)
                    {
                        // Create new URI
                        int dot = part.Uri.ToString().LastIndexOf(".");
                        string new_path = part.Uri.ToString().Substring(0, dot) + ".tiff";
                        Uri new_uri = new Uri(new_path, UriKind.Relative);

                        // Convert data
                        Stream stream = part.GetStream();
                        MemoryStream new_stream = new MemoryStream();
                        using (MagickImage image = new MagickImage(stream))
                        {
                            // Set input stream position to beginning
                            stream.Position = 0;

                            // Adjust TIFF settings
                            image.Format = MagickFormat.Tiff;
                            image.Settings.ColorSpace = ColorSpace.RGB;
                            image.Settings.Depth = 32;
                            image.Settings.Compression = CompressionMethod.LZW;

                            // Write image to stream
                            image.Write(new_stream);
                        }
                        stream.Dispose();

                        // Add new ImagePart
                        ImagePart new_ImagePart = worksheetPart.VmlDrawingParts.First().AddImagePart(ImagePartType.Tiff);

                        // Save image from stream to new ImagePart
                        new_stream.Position = 0;
                        new_ImagePart.FeedData(new_stream);

                        // Change relationships of image
                        string id = GetRelationshipId(part);
                        ImageData imageData = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<ImageData>()
                                        .Where(p => p.RelId == id)
                                        .Select(p => p)
                                        .Single();
                        imageData.RelId = GetRelationshipId(new_ImagePart);

                        // Delete original ImagePart
                        worksheetPart.VmlDrawingParts.First().DeletePart(part);
                    }
                    foreach (ImagePart part in images)
                    {
                        // Create new URI
                        int dot = part.Uri.ToString().LastIndexOf(".");
                        string new_path = part.Uri.ToString().Substring(0, dot) + ".tiff";
                        Uri new_uri = new Uri(new_path, UriKind.Relative);

                        // Convert data
                        Stream stream = part.GetStream();
                        MemoryStream new_stream = new MemoryStream();
                        using (MagickImage image = new MagickImage(stream))
                        {
                            // Set input stream position to beginning
                            stream.Position = 0;

                            // Adjust TIFF settings
                            image.Format = MagickFormat.Tiff;
                            image.Settings.ColorSpace = ColorSpace.RGB;
                            image.Settings.Depth = 32;
                            image.Settings.Compression = CompressionMethod.LZW;

                            // Write image to stream
                            image.Write(new_stream);
                        }
                        stream.Dispose();

                        // Add new ImagePart
                        ImagePart new_ImagePart = worksheetPart.DrawingsPart.AddImagePart(ImagePartType.Tiff);

                        // Save image from stream to new ImagePart
                        new_stream.Position = 0;
                        new_ImagePart.FeedData(new_stream);

                        // Change relationships of image
                        string id = GetRelationshipId(part);
                        Blip blip = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                                        .Where(p => p.BlipFill.Blip.Embed == id)
                                        .Select(p => p.BlipFill.Blip)
                                        .Single();
                        blip.Embed = GetRelationshipId(new_ImagePart);

                        // Delete original ImagePart
                        worksheetPart.DrawingsPart.DeletePart(part);
                    }
                }
            }
        }

        // Get relationship id of an OpenXmlPart. Is used by other methods.
        private string GetRelationshipId(OpenXmlPart part)
        {
            string id = "";
            IEnumerable<OpenXmlPart> parentParts = part.GetParentParts();
            foreach (OpenXmlPart parentPart in parentParts)
            {
                if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.DrawingsPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.VmlDrawingPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.Model3DReferenceRelationshipPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
                else if (parentPart.ToString() == "DocumentFormat.OpenXml.Packaging.OleObjectPart")
                {
                    id = parentPart.GetIdOfPart(part);
                    return id;
                }
            }
            return id;
        }
    }
}
