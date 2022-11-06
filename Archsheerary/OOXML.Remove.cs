using System;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2013.ExcelAc;

namespace Archsheerary
{
    public partial class OOXML
    {
        public class Remove
        {
            // Remove data connections
            public List<Lists.DataConnections> DataConnections(string filepath)
            {
                List<Lists.DataConnections> results = new List<Lists.DataConnections>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    // Find all connections
                    ConnectionsPart conns = spreadsheet.WorkbookPart.ConnectionsPart;

                    // Write information to list
                    foreach (Connection conn in conns.Connections)
                    {
                        results.Add(new Lists.DataConnections() { Id = conn.Id, Description = conn.Description, ConnectionFile = conn.ConnectionFile, Credentials = conn.Credentials, DatabaseProperties = conn.DatabaseProperties.ToString(), Action = Lists.ActionRemoved });
                    }

                    // Delete connections
                    spreadsheet.WorkbookPart.DeletePart(conns);

                    // Delete all query tables
                    List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    foreach (WorksheetPart part in worksheetparts)
                    {
                        List<QueryTablePart> queryTables = part.QueryTableParts.ToList();
                        foreach (QueryTablePart qtp in queryTables)
                        {
                            part.DeletePart(qtp);
                        }
                    }

                    // If spreadsheet contains a custom XML Map, delete databinding
                    if (spreadsheet.WorkbookPart.CustomXmlMappingsPart != null)
                    {
                        CustomXmlMappingsPart xmlMap = spreadsheet.WorkbookPart.CustomXmlMappingsPart;
                        List<Map> maps = xmlMap.MapInfo.Elements<Map>().ToList();
                        foreach (Map map in maps)
                        {
                            if (map.DataBinding != null)
                            {
                                map.DataBinding.Remove();
                            }
                        }
                    }
                }
                // Repair spreadsheet
                //Repair rep = new Repair();
                //rep.Repair_QueryTables(filepath);

                return results;
            }

            // Remove RTD functions
            public List<Lists.RTDFunctions> RTDFunctions(string filepath)
            {
                List<Lists.RTDFunctions> results = new List<Lists.RTDFunctions>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    foreach (WorksheetPart part in worksheetparts)
                    {
                        Worksheet worksheet = part.Worksheet;
                        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                        foreach (var row in rows)
                        {
                            var cells = row.Elements<Cell>();
                            foreach (Cell cell in cells)
                            {
                                if (cell.CellFormula != null)
                                {
                                    string formula = cell.CellFormula.InnerText;
                                    if (formula.Length > 2)
                                    {
                                        string hit = formula.Substring(0, 3); // Transfer first 3 characters to string
                                        if (hit == "RTD")
                                        {
                                            // Add to list
                                            results.Add(new Lists.RTDFunctions() { Sheet = worksheet.NamespaceUri, Cell = cell.CellReference, Value = cell.CellValue.ToString(), Formula = cell.CellFormula.ToString(), Action = Lists.ActionRemoved });
                                            
                                            // Remove
                                            CellValue cellvalue = cell.CellValue; // Save current cell value
                                            cell.CellFormula = null; // Remove RTD formula
                                            // If cellvalue does not have a real value
                                            if (cellvalue.Text == "#N/A")
                                            {
                                                cell.DataType = CellValues.String;
                                                cell.CellValue = new CellValue("Invalid data removed");
                                            }
                                            else
                                            {
                                                cell.CellValue = cellvalue; // Insert saved cell value
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Delete calculation chain
                    CalculationChainPart calc = spreadsheet.WorkbookPart.CalculationChainPart;
                    spreadsheet.WorkbookPart.DeletePart(calc);

                    // Delete volatile dependencies
                    VolatileDependenciesPart vol = spreadsheet.WorkbookPart.VolatileDependenciesPart;
                    spreadsheet.WorkbookPart.DeletePart(vol);
                }
                return results;
            }

            // Remove printer settings
            public List<Lists.PrinterSettings> PrinterSettings(string filepath)
            {
                List<Lists.PrinterSettings> results = new List<Lists.PrinterSettings>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    List<WorksheetPart> wsParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    foreach (WorksheetPart wsPart in wsParts)
                    {
                        List<SpreadsheetPrinterSettingsPart> printerList = wsPart.SpreadsheetPrinterSettingsParts.ToList();
                        foreach (SpreadsheetPrinterSettingsPart printer in printerList)
                        {
                            // Add to list
                            results.Add(new Lists.PrinterSettings() { Uri = printer.Uri.ToString(), Action = Lists.ActionRemoved });

                            // Delete printer
                            wsPart.DeletePart(printer);
                        }
                    }
                }
                return results;
            }

            // Remove external cell references
            public List<Lists.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<Lists.ExternalCellReferences> results = new List<Lists.ExternalCellReferences>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    foreach (WorksheetPart part in worksheetparts)
                    {
                        Worksheet worksheet = part.Worksheet;
                        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                        foreach (var row in rows)
                        {
                            var cells = row.Elements<Cell>();
                            foreach (Cell cell in cells)
                            {
                                if (cell.CellFormula != null)
                                {
                                    string formula = cell.CellFormula.InnerText;
                                    if (formula.Length > 1)
                                    {
                                        string hit = formula.Substring(0, 1); // Transfer first 1 characters to string
                                        string hit2 = formula.Substring(0, 2); // Transfer first 2 characters to string
                                        if (hit == "[" || hit2 == "'[")
                                        {
                                            // Add to list
                                            results.Add(new Lists.ExternalCellReferences() { Sheet = worksheet.NamespaceUri, Cell = cell.CellReference, Value = cell.CellValue.ToString(), Formula = cell.CellFormula.ToString(), Action = Lists.ActionRemoved });

                                            // Remove
                                            CellValue cellvalue = cell.CellValue; // Save current cell value
                                            cell.CellFormula = null;
                                            // If cellvalue does not have a real value
                                            if (cellvalue.Text == "#N/A")
                                            {
                                                cell.DataType = CellValues.String;
                                                cell.CellValue = new CellValue("Invalid data removed");
                                            }
                                            else
                                            {
                                                cell.CellValue = cellvalue; // Insert saved cell value
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Delete external book references
                    List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
                    if (extwbParts.Count > 0)
                    {
                        foreach (ExternalWorkbookPart extpart in extwbParts)
                        {
                            var elements = extpart.ExternalLink.ChildElements.ToList();
                            foreach (var element in elements)
                            {
                                if (element.LocalName == "externalBook")
                                {
                                    spreadsheet.WorkbookPart.DeletePart(extpart);
                                }
                            }
                        }
                    }

                    // Delete calculation chain
                    CalculationChainPart calc = spreadsheet.WorkbookPart.CalculationChainPart;
                    spreadsheet.WorkbookPart.DeletePart(calc);

                    // Delete defined names that includes external cell references
                    DefinedNames definedNames = spreadsheet.WorkbookPart.Workbook.DefinedNames;
                    if (definedNames != null)
                    {
                        var definedNamesList = definedNames.ToList();
                        foreach (DefinedName definedName in definedNamesList)
                        {
                            if (definedName.InnerXml.StartsWith("["))
                            {
                                definedName.Remove();
                            }
                        }
                    }
                }
                return results;
            }

            // Remove external object references
            public List<Lists.ExternalObjects> ExternalObjects(string filepath)
            {
                List<Lists.ExternalObjects> results = new List<Lists.ExternalObjects>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                    foreach (ExternalWorkbookPart extWbPart in extWbParts)
                    {
                        List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList();
                        foreach (ExternalRelationship extrel in extrels)
                        {
                            // Add to list
                            results.Add(new Lists.ExternalObjects() { Uri = extrel.Uri.ToString(), Target = extrel., IsExternal = extrel.IsExternal, Action = Lists.ActionRemoved });

                            // Change external target reference
                            Uri uri = new Uri("External reference was removed", UriKind.Relative);
                            extWbPart.DeleteExternalRelationship("rId1");
                            extWbPart.AddExternalRelationship(relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject", externalUri: uri, id: "rId1");
                        }
                    }
                }
                return results;
            }

            public List<Lists.EmbeddedObjects> EmbeddedObjects(string filepath)
            {
                List<Lists.EmbeddedObjects> results = new List<Lists.EmbeddedObjects>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        List<EmbeddedObjectPart> embedobj_ole_list = worksheetPart.EmbeddedObjectParts.ToList();
                        List<EmbeddedPackagePart> embedobj_package_list = worksheetPart.EmbeddedPackageParts.ToList();
                        List<ImagePart> embedobj_image_list = worksheetPart.ImageParts.ToList();
                        List<ImagePart> embedobj_drawing_image_list = new List<ImagePart>();
                        if (worksheetPart.DrawingsPart != null)
                        {
                            embedobj_drawing_image_list = worksheetPart.DrawingsPart.ImageParts.ToList();
                        }
                        List<Model3DReferenceRelationshipPart> embedobj_3d_list = worksheetPart.Model3DReferenceRelationshipParts.ToList();

                        if (embedobj_ole_list.Count() > 0)
                        {
                            foreach (EmbeddedObjectPart ole in embedobj_ole_list)
                            {
                                worksheetPart.DeletePart(ole);
                            }
                        }

                        if (embedobj_package_list.Count() > 0)
                        {
                            foreach (EmbeddedPackagePart package in embedobj_package_list)
                            {
                                worksheetPart.DeletePart(package);
                            }
                        }
                        if (embedobj_image_list.Count() > 0)
                        {
                            foreach (ImagePart image in embedobj_image_list)
                            {
                                worksheetPart.DeletePart(image);
                            }
                        }
                        if (embedobj_drawing_image_list.Count() > 0)
                        {
                            foreach (ImagePart drawing_image in embedobj_drawing_image_list)
                            {
                                worksheetPart.DrawingsPart.DeletePart(drawing_image);
                            }
                        }
                        if (embedobj_3d_list.Count() > 0)
                        {
                            foreach (Model3DReferenceRelationshipPart threeD in embedobj_3d_list)
                            {
                                worksheetPart.DeletePart(threeD);
                            }
                        }
                    }
                }
                return results;
            }

            // Remove absolute path to local directory
            public List<Lists.AbsolutePath> AbsolutePath(string filepath)
            {
                List<Lists.AbsolutePath> results = new List<Lists.AbsolutePath>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                    {
                        AbsolutePath absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();

                        // Add to list
                        results.Add(new Lists.AbsolutePath() { Path = absPath.ToString(), Action = Lists.ActionRemoved });

                        // Remove
                        absPath.Remove();
                    }
                }
                return results;
            }

            // Remove metadata in file properties
            public void FilePropertyInformation(string input_filepath, string output_folder)
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, true))
                {
                    PackageProperties property = spreadsheet.Package.PackageProperties;

                    // Create metadata file
                    using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_Metadata.txt"))
                    {
                        w.WriteLine("STRIPPED FILE PROPERTIES INFORMATION");
                        w.WriteLine("---");
                    }

                    if (property.Creator != null)
                    {
                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"CREATOR: {property.Creator}");
                        }

                        // Remove information
                        property.Creator = null;
                    }
                    if (property.Title != null)
                    {
                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"TITLE: {property.Title}");
                        }

                        // Remove information
                        property.Title = null;
                    }
                    if (property.Subject != null)
                    {
                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"SUBJECT: {property.Subject}");
                        }

                        // Remove information
                        property.Subject = null;
                    }
                    if (property.Description != null)
                    {
                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"DESCRIPTION: {property.Description}");
                        }

                        // Remove information
                        property.Description = null;
                    }
                    if (property.Keywords != null)
                    {

                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"KEYWORDS: {property.Keywords}");
                        }

                        // Remove information
                        property.Keywords = null;
                    }
                    if (property.Category != null)
                    {
                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"CATEGORY: {property.Category}");
                        }

                        // Remove information
                        property.Category = null;
                    }
                    if (property.LastModifiedBy != null)
                    {
                        // Write information to metadata file
                        using (StreamWriter w = File.AppendText($"{output_folder}\\orgFile_metadata.txt"))
                        {
                            w.WriteLine($"LAST MODIFIED BY: {property.LastModifiedBy}");
                        }

                        // Remove information
                        property.LastModifiedBy = null;
                    }
                }
            }
        }
    }
}
