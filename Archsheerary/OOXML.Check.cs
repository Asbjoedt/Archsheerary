using System;
using System.IO;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2013.ExcelAc;

namespace Archsheerary
{
    public partial class OOXML
    {
        /// <summary>
        /// Collection of methods for checking content in Office Open XML spreadsheets.
        /// </summary>
        public class Check
        {
            /// <summary>
            /// Check for existence of any cell values.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>True if cell values were identified</returns>
            public static bool ValuesExist(string filepath)
            {
                bool hascellvalues = false;

                // Perform check
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    if (spreadsheet.WorkbookPart.WorksheetParts != null)
                    {
                        List<WorksheetPart> worksheetparts = spreadsheet.WorkbookPart.WorksheetParts.ToList();
                        foreach (WorksheetPart part in worksheetparts)
                        {
                            Worksheet worksheet = part.Worksheet;
                            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                            if (rows.Count() > 0) // If any rows exist, this means cells exist
                            {
                                hascellvalues = true;
                            }
                        }
                    }
                }
                return hascellvalues;
            }

            /// <summary>
            /// Check for conformance of XLSX file.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified conformance</returns>
            public static List<DataTypes.Conformance> Conformance(string filepath)
            {
                List<DataTypes.Conformance> results = new List<DataTypes.Conformance>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                    if (workbook.Conformance == null || workbook.Conformance == "transitional")
                    {
                        // Add to list
                        results.Add(new DataTypes.Conformance() { OriginalConformance = "Transitional", NewConformance = null, Action = DataTypes.ActionChecked });
                    }
                    else if (workbook.Conformance == "strict")
                    {
                        // Add to list
                        results.Add(new DataTypes.Conformance() { OriginalConformance = "Strict", NewConformance = null, Action = DataTypes.ActionChecked });
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for data connections.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <return>List of identified data connections</return>
            public static List<DataTypes.DataConnections> DataConnections(string filepath)
            {
                List<DataTypes.DataConnections> results = new List<DataTypes.DataConnections>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    ConnectionsPart conns = spreadsheet.WorkbookPart.ConnectionsPart;
                    if (conns != null)
                    {
                        // Write information to list
                        foreach (Connection conn in conns.Connections)
                        {
                            results.Add(new DataTypes.DataConnections() { Id = conn.Id, Name = conn.Name, Description = conn.Description, Type = conn.Type, ConnectionFile = conn.ConnectionFile, Credentials = conn.Credentials, DatabaseProperties = conn.DatabaseProperties.ToString(), SourceFile = conn. SourceFile, Action = DataTypes.ActionChecked });
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for external cell references.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified external cell references</returns>
            public static List<DataTypes.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<DataTypes.ExternalCellReferences> results = new List<DataTypes.ExternalCellReferences>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                    foreach (WorksheetPart part in worksheetParts)
                    {
                        Worksheet worksheet = part.Worksheet;
                        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                        foreach (var row in rows)
                        {
                            IEnumerable<Cell> cells = row.Elements<Cell>();
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
                                            results.Add(new DataTypes.ExternalCellReferences() { Sheet = worksheet.NamespaceUri, Cell = cell.CellReference, Value = cell.CellValue.ToString(), Formula = cell.CellFormula.ToString(), Action = DataTypes.ActionChecked });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for external object references.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified external object references</returns>
            public static List<DataTypes.ExternalObjects> ExternalObjects(string filepath)
            {
                List<DataTypes.ExternalObjects> results = new List<DataTypes.ExternalObjects>();

                // Perform check
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                    foreach (ExternalWorkbookPart extWbPart in extWbParts)
                    {
                        List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList();
                        foreach (ExternalRelationship extrel in extrels)
                        {
                            results.Add(new DataTypes.ExternalObjects() { Target = extrel.Uri.ToString(), RelationshipType = extrel.RelationshipType, IsExternal = extrel.IsExternal, Container = extrel.Container.ToString(), Action = DataTypes.ActionChecked });
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for RealTimeData (RTD) functions.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified RTD functions</returns>
            public static List<DataTypes.RTDFunctions> RTDFunctions(string filepath) // Check for RTD functions
            {
                List<DataTypes.RTDFunctions> results = new List<DataTypes.RTDFunctions>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                    foreach (WorksheetPart part in worksheetParts)
                    {
                        Worksheet worksheet = part.Worksheet;
                        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
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
                                            results.Add(new DataTypes.RTDFunctions() { Sheet = worksheet.NamespaceUri, Cell = cell.CellReference, Value = cell.CellValue.ToString(), Formula = cell.CellFormula.ToString(), Action = DataTypes.ActionChecked });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for embedded objects.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified embedded objects</returns>
            public static List<DataTypes.EmbeddedObjects> EmbeddedObjects(string filepath)
            {
                List<DataTypes.EmbeddedObjects> results = new List<DataTypes.EmbeddedObjects>();
                List<EmbeddedObjectPart> embeddings_ole = new List<EmbeddedObjectPart>();
                List<EmbeddedPackagePart> embeddings_package = new List<EmbeddedPackagePart>();
                List<ImagePart> embeddings_emf = new List<ImagePart>();
                List<ImagePart> embeddings_image = new List<ImagePart>();
                List<Model3DReferenceRelationshipPart> embeddings_3d = new List<Model3DReferenceRelationshipPart>();

                using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;

                    // Perform check
                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        embeddings_ole = worksheetPart.EmbeddedObjectParts.Distinct().ToList();
                        embeddings_package = worksheetPart.EmbeddedPackageParts.Distinct().ToList();
                        embeddings_emf = worksheetPart.ImageParts.Distinct().ToList();
                        if (worksheetPart.DrawingsPart != null) // DrawingsPart needs a null check
                        {
                            embeddings_image = worksheetPart.DrawingsPart.ImageParts.Distinct().ToList();
                        }
                        embeddings_3d = worksheetPart.Model3DReferenceRelationshipParts.Distinct().ToList();
                    }

                    if (embeddings_ole.Count() > 0)
                    {
                        foreach (EmbeddedObjectPart part in embeddings_ole)
                        {
                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionChecked });
                        }
                    }
                    if (embeddings_package.Count() > 0)
                    {
                        foreach (EmbeddedPackagePart part in embeddings_package)
                        {
                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionChecked });
                        }
                    }
                    if (embeddings_emf.Count() > 0)
                    {
                        foreach (ImagePart part in embeddings_emf)
                        {
                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionChecked });
                        }
                    }
                    if (embeddings_image.Count() > 0)
                    {
                        foreach (ImagePart part in embeddings_image)
                        {
                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionChecked });
                        }
                    }
                    if (embeddings_3d.Count() > 0)
                    {
                        foreach (Model3DReferenceRelationshipPart part in embeddings_3d)
                        {
                            // Add to list
                            results.Add(new DataTypes.EmbeddedObjects() { Uri = part.Uri.ToString(), ContentType = part.ContentType, RelationshipType = part.RelationshipType, Action = DataTypes.ActionChecked });
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for hyperlinks.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified hyperlinks</returns>
            public static List<DataTypes.Hyperlinks> Hyperlinks(string filepath)
            {
                List<DataTypes.Hyperlinks> results = new List<DataTypes.Hyperlinks>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    List<HyperlinkRelationship> hyperlinks = spreadsheet
                        .GetAllParts()
                        .SelectMany(p => p.HyperlinkRelationships)
                        .ToList();

                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        // Add to list
                        results.Add(new DataTypes.Hyperlinks() { URL = hyperlink.Uri.ToString(), Action = DataTypes.ActionChecked });
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for printer settings.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified printer settings</returns>
            public static List<DataTypes.PrinterSettings> PrinterSettings(string filepath)
            {
                List<DataTypes.PrinterSettings> results = new List<DataTypes.PrinterSettings>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    IEnumerable<WorksheetPart> worksheetParts = spreadsheet.WorkbookPart.WorksheetParts;
                    List<SpreadsheetPrinterSettingsPart> printerList = new List<SpreadsheetPrinterSettingsPart>();
                    foreach (WorksheetPart worksheetPart in worksheetParts)
                    {
                        printerList = worksheetPart.SpreadsheetPrinterSettingsParts.ToList();
                    }
                    foreach (SpreadsheetPrinterSettingsPart printer in printerList)
                    {
                        // Add to list
                        results.Add(new DataTypes.PrinterSettings() { Uri = printer.Uri.ToString(), Action = DataTypes.ActionChecked });
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for active sheet.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified active sheet</returns>
            public static List<DataTypes.ActiveSheet> ActiveSheet(string filepath)
            {
                List<DataTypes.ActiveSheet> results = new List<DataTypes.ActiveSheet>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    if (spreadsheet.WorkbookPart.Workbook.BookViews != null)
                    {
                        BookViews bookViews = spreadsheet.WorkbookPart.Workbook.BookViews;
                        if (bookViews.ChildElements.Where(p => p.OuterXml == "workbookView") != null)
                        {
                            WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                            if (workbookView.ActiveTab != null)
                            {

                                results.Add(new DataTypes.ActiveSheet() { OriginalActiveSheet = workbookView.ActiveTab.Value, NewActiveSheet = null, Action = DataTypes.ActionChecked });
                            }
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Check for absolute path to local directory.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified absolute path</returns>
            public static List<DataTypes.AbsolutePath> AbsolutePath(string filepath)
            {
                List<DataTypes.AbsolutePath> results = new List<DataTypes.AbsolutePath>();
                AbsolutePath absPath = null;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                    {
                        absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();
                    }
                    // Add to list
                    results.Add(new DataTypes.AbsolutePath() { Path = absPath.ToString(), Action = DataTypes.ActionChecked });
                }
                return results;
            }

            /// <summary>
            /// Check for file property information.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <return>List of identified file property information</return>
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
                bool found = false;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    PackageProperties property = spreadsheet.Package.PackageProperties;

                    if (property.Creator != null)
                    {
                        creator = property.Creator;
                        found = true;
                    }
                    if (property.Title != null)
                    {
                        title = property.Title;
                        found = true;
                    }
                    if (property.Subject != null)
                    {
                        subject = property.Subject;
                        found = true;
                    }
                    if (property.Description != null)
                    {
                        description = property.Description;
                        found = true;
                    }
                    if (property.Keywords != null)
                    {
                        keywords = property.Keywords;
                        found = true;
                    }
                    if (property.Category != null)
                    {
                        category = property.Category;
                        found = true;
                    }
                    if (property.LastModifiedBy != null)
                    {
                        lastmodifiedby = property.LastModifiedBy;
                        found = true;
                    }

                    // Add to list
                    results.Add(new DataTypes.FilePropertyInformation() { Author = creator, Title = title, Subject = subject, Description = description, Keywords = keywords, Category = category, LastModifiedBy = lastmodifiedby, FilePropertyInfoFound = found, Action = DataTypes.ActionChecked });
                }
                return results;
            }
        }
    }
}
