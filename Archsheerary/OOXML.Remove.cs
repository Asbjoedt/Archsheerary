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
        /// <summary>
        /// Collection of methods for removing content in Office Open XML spreadsheets.
        /// </summary>
        public class Remove
        {
            /// <summary>
            /// Remove data connections.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed data connections</returns>
            public static List<DataTypes.DataConnections> DataConnections(string filepath)
            {
                List<DataTypes.DataConnections> results = new List<DataTypes.DataConnections>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    // Find all connections
                    ConnectionsPart conns = spreadsheet.WorkbookPart.ConnectionsPart;

                    // Write information to list
                    foreach (Connection conn in conns.Connections)
                    {
                        results.Add(new DataTypes.DataConnections() { Id = conn.Id, Description = conn.Description, ConnectionFile = conn.ConnectionFile, Credentials = conn.Credentials, DatabaseProperties = conn.DatabaseProperties.ToString(), Action = DataTypes.ActionRemoved });
                    }

                    // Delete connections
                    spreadsheet.WorkbookPart.DeletePart(conns);

                    // Delete all QueryTableParts
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        // Delete all QueryTableParts in WorksheetParts
                        List<QueryTablePart> queryTables = worksheetPart.QueryTableParts.ToList(); // Must be a list
                        foreach (QueryTablePart queryTablePart in queryTables)
                        {
                            worksheetPart.DeletePart(queryTablePart);
                        }

                        // Delete all QueryTableParts, if they are not registered in a WorksheetPart
                        List<TableDefinitionPart> tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
                        foreach (TableDefinitionPart tableDefinitionPart in tableDefinitionParts)
                        {
                            List<IdPartPair> idPartPairs = tableDefinitionPart.Parts.ToList();
                            foreach (IdPartPair idPartPair in idPartPairs)
                            {
                                if (idPartPair.OpenXmlPart.ToString() == "DocumentFormat.OpenXml.Packaging.QueryTablePart")
                                {
                                    // Delete QueryTablePart
                                    tableDefinitionPart.DeletePart(idPartPair.OpenXmlPart);
                                    // The TableDefinitionPart must also be deleted
                                    worksheetPart.DeletePart(tableDefinitionPart);
                                    // And the reference to the TableDefinitionPart in the WorksheetPart must be deleted
                                    List<TablePart> tableParts = worksheetPart.Worksheet.Descendants<TablePart>().ToList();
                                    foreach (TablePart tablePart in tableParts)
                                    {
                                        if (idPartPair.RelationshipId == tablePart.Id)
                                        {
                                            tablePart.Remove();
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // If spreadsheet contains a CustomXmlMappingsPart, delete databinding
                    if (spreadsheet.WorkbookPart.CustomXmlMappingsPart != null)
                    {
                        CustomXmlMappingsPart xmlMap = spreadsheet.WorkbookPart.CustomXmlMappingsPart;
                        List<Map> maps = xmlMap.MapInfo.Elements<Map>().ToList(); // Must be a list
                        foreach (Map map in maps)
                        {
                            if (map.DataBinding != null)
                            {
                                map.DataBinding.Remove();
                            }
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Remove RealTimeData (RTD) functions.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed RTD functions</returns>
            public static List<DataTypes.RTDFunctions> RTDFunctions(string filepath)
            {
                List<DataTypes.RTDFunctions> results = new List<DataTypes.RTDFunctions>();

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
                                            results.Add(new DataTypes.RTDFunctions() { Sheet = worksheet.NamespaceUri, Cell = cell.CellReference, Value = cell.CellValue.ToString(), Formula = cell.CellFormula.ToString(), Action = DataTypes.ActionRemoved });
                                            
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

            /// <summary>
            /// Remove printer settings.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed removed printer settings</returns>
            public static List<DataTypes.PrinterSettings> PrinterSettings(string filepath)
            {
                List<DataTypes.PrinterSettings> results = new List<DataTypes.PrinterSettings>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    List<WorksheetPart> wsParts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                    foreach (WorksheetPart wsPart in wsParts)
                    {
                        List<SpreadsheetPrinterSettingsPart> printerList = wsPart.SpreadsheetPrinterSettingsParts.ToList();
                        foreach (SpreadsheetPrinterSettingsPart printer in printerList)
                        {
                            // Add to list
                            results.Add(new DataTypes.PrinterSettings() { Uri = printer.Uri.ToString(), Action = DataTypes.ActionRemoved });

                            // Delete printer
                            wsPart.DeletePart(printer);
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Remove external cell references.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed external cell references</returns>
            public static List<DataTypes.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<DataTypes.ExternalCellReferences> results = new List<DataTypes.ExternalCellReferences>();

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
                                            results.Add(new DataTypes.ExternalCellReferences() { Sheet = worksheet.NamespaceUri, Cell = cell.CellReference, Value = cell.CellValue.ToString(), Formula = cell.CellFormula.ToString(), Action = DataTypes.ActionRemoved });

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

            /// <summary>
            /// Remove external object references.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed external object references</returns>
            public static List<DataTypes.ExternalObjects> ExternalObjects(string filepath)
            {
                List<DataTypes.ExternalObjects> results = new List<DataTypes.ExternalObjects>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                    foreach (ExternalWorkbookPart extWbPart in extWbParts)
                    {
                        List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList();
                        foreach (ExternalRelationship extrel in extrels)
                        {
                            // Add to list
                            results.Add(new DataTypes.ExternalObjects() { Target = extrel.Uri.ToString(), RelationshipType = extrel.RelationshipType, IsExternal = extrel.IsExternal, Container = extrel.Container.ToString(), Action = DataTypes.ActionRemoved });

                            // Change external target reference
                            Uri uri = new Uri($"External reference {extrel.Uri} was removed", UriKind.Relative);
                            extWbPart.DeleteExternalRelationship(extrel.Id);
                            extWbPart.AddExternalRelationship(relationshipType: "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject", externalUri: uri, id: extrel.Id);
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Remove embedded objects.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed embedded objects</returns>
            public static List<DataTypes.EmbeddedObjects> EmbeddedObjects(string filepath)
            {
                List<DataTypes.EmbeddedObjects> results = new List<DataTypes.EmbeddedObjects>();

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
                            foreach (EmbeddedObjectPart part in embedobj_ole_list)
                            {
                                // Add to list
                                results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionRemoved });

                                //Remove
                                worksheetPart.DeletePart(part);
                            }
                        }

                        if (embedobj_package_list.Count() > 0)
                        {
                            foreach (EmbeddedPackagePart part in embedobj_package_list)
                            {
                                // Add to list
                                results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionRemoved });

                                //Remove
                                worksheetPart.DeletePart(part);
                            }
                        }
                        if (embedobj_image_list.Count() > 0)
                        {
                            foreach (ImagePart part in embedobj_image_list)
                            {
                                // Add to list
                                results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionRemoved });

                                //Remove
                                worksheetPart.DeletePart(part);
                            }
                        }
                        if (embedobj_drawing_image_list.Count() > 0)
                        {
                            foreach (ImagePart part in embedobj_drawing_image_list)
                            {
                                // Add to list
                                results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionRemoved });

                                //Remove
                                worksheetPart.DeletePart(part);
                            }
                        }
                        if (embedobj_3d_list.Count() > 0)
                        {
                            foreach (Model3DReferenceRelationshipPart part in embedobj_3d_list)
                            {
                                // Add to list
                                results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionRemoved });

                                //Remove
                                worksheetPart.DeletePart(part);
                            }
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Remove absolute path to local directory.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed absolute path</returns>
            public static List<DataTypes.AbsolutePath> AbsolutePath(string filepath)
            {
                List<DataTypes.AbsolutePath> results = new List<DataTypes.AbsolutePath>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                    {
                        AbsolutePath absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();

                        // Add to list
                        results.Add(new DataTypes.AbsolutePath() { Path = absPath.ToString(), Action = DataTypes.ActionRemoved });

                        // Remove
                        absPath.Remove();
                    }
                }
                return results;
            }

            /// <summary>
            /// Remove file property information.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of removed file property information</returns>
            public static List<DataTypes.FilePropertyInformation> FilePropertyInformation(string filepath)
            {
                List<DataTypes.FilePropertyInformation> results = new List<DataTypes.FilePropertyInformation>();
                string creator = "";
                string title = "";
                string subject = "";
                string description = "";
                string keywords = "";
                string category = "";
                string lastmodifiedby = "";

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    PackageProperties property = spreadsheet.Package.PackageProperties;

                    if (property.Creator != null)
                    {
                        creator = property.Creator;

                        // Remove information
                        property.Creator = null;
                    }
                    if (property.Title != null)
                    {
                        title = property.Title;

                        // Remove information
                        property.Title = null;
                    }
                    if (property.Subject != null)
                    {
                        subject = property.Subject;

                        // Remove information
                        property.Subject = null;
                    }
                    if (property.Description != null)
                    {
                        description = property.Description;

                        // Remove information
                        property.Description = null;
                    }
                    if (property.Keywords != null)
                    {
                        keywords = property.Keywords;

                        // Remove information
                        property.Keywords = null;
                    }
                    if (property.Category != null)
                    {
                        category = property.Category;

                        // Remove information
                        property.Category = null;
                    }
                    if (property.LastModifiedBy != null)
                    {
                        lastmodifiedby = property.LastModifiedBy;

                        // Remove information
                        property.LastModifiedBy = null;
                    }

                    // Add to list
                    results.Add(new DataTypes.FilePropertyInformation() { Author = creator, Title = title, Subject = subject, Description = description, Keywords = keywords, Category = category, LastModifiedBy = lastmodifiedby, Action = DataTypes.ActionRemoved });
                }
                return results;
            }
        }
    }
}
