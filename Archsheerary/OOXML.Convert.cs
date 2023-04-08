using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Pictures;
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
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_filepath">Path to output file</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToXLSXTransitional(string input_filepath, string output_filepath)
            {
                bool convert_success = false;

                // Convert spreadsheet
                byte[] byteArray = File.ReadAllBytes(input_filepath);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                    {
                        // Perform conversion
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
        /// <param name="filepath">Path to input file</param>
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
                        XDocument xElement = worksheetPart.VmlDrawingParts.First().GetXDocument();
                        IEnumerable<XElement> descendants = xElement.FirstNode.Document.Descendants();
                        foreach (XElement descendant in descendants)
                        {
                            if (descendant.Name == "{urn:schemas-microsoft-com:vml}imagedata")
                            {
                                IEnumerable<XAttribute> attributes = descendant.Attributes();
                                foreach (XAttribute attribute in attributes)
                                {
                                    if (attribute.Name == "{urn:schemas-microsoft-com:office:office}relid")
                                    {
                                        if (attribute.Value == id)
                                        {
                                            attribute.Value = GetRelationshipId(new_ImagePart);
                                            worksheetPart.VmlDrawingParts.First().SaveXDocument();
                                        }
                                    }
                                }
                            }
                        }

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

        // Get relationship id of an OpenXmlPart. Is used by other methods. You may ignore it.
        internal string GetRelationshipId(OpenXmlPart part)
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
