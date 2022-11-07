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
        public class Check
        {
            // Check for any values by checking if sheets and cell values exist
            public bool ValuesExist(string filepath)
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

            // Check for Strict conformance
            public bool Conformance(string filepath)
            {
                bool conformance = false;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                    if (workbook.Conformance == null || workbook.Conformance == "transitional")
                    {
                        conformance = true;
                    }
                    else if (workbook.Conformance == "strict")
                    {
                        conformance = false;
                    }
                }
                return conformance;
            }

            // Check for data connections
            public int DataConnections(string filepath)
            {
                int conn_count = 0;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
                    if (conn != null)
                    {
                        conn_count = conn.Connections.Count();
                    }
                }
                return conn_count;
            }

            // Check for external cell references
            public int ExternalCellReferences(string filepath)
            {
                int ext_cellrefs_count = 0;

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
                                            ext_cellrefs_count++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return ext_cellrefs_count;
            }

            // Check for external object references
            public List<Lists.ExternalObjects> ExternalObjects(string filepath)
            {
                List<Lists.ExternalObjects> results = new List<Lists.ExternalObjects>();

                // Perform check
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    IEnumerable<ExternalWorkbookPart> extWbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts;
                    foreach (ExternalWorkbookPart extWbPart in extWbParts)
                    {
                        List<ExternalRelationship> extrels = extWbPart.ExternalRelationships.ToList();
                        foreach (ExternalRelationship extrel in extrels)
                        {
                            results.Add(new Lists.ExternalObjects() { Uri = extrel.Uri.ToString(), Target = extrel., IsExternal = extrel.IsExternal.ToString(), Action = Lists.ActionChecked });
                        }
                    }
                }
                return results;
            }

            // Check for RTD functions
            public static int RTDFunctions(string filepath) // Check for RTD functions
            {
                int rtd_functions_count = 0;

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
                                            rtd_functions_count++;
                                            Console.WriteLine($"--> Check: RTD function in sheet {part.Uri} cell {cell.CellReference} detected");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return rtd_functions_count;
            }

            // Check for embedded objects
            public int EmbeddedObjects(string filepath)
            {
                int count_embedobj = 0;
                int embedobj_number = 0;
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

                    // Count number of embeddings
                    count_embedobj = embeddings_ole.Count() + embeddings_package.Count() + embeddings_emf.Count() + embeddings_image.Count() + embeddings_3d.Count();

                    // Inform user of detected embedded objects
                    if (count_embedobj > 0)
                    {
                        foreach (EmbeddedObjectPart part in embeddings_ole)
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: OLE object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        // Inform user of each package object
                        foreach (EmbeddedPackagePart part in embeddings_package)
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: Package object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        // Inform user of each .emf image object
                        foreach (ImagePart part in embeddings_emf)
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: Rendering (.emf) of embeddings object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        // Inform user of each image object
                        foreach (ImagePart part in embeddings_image)
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: Image object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                        // Inform user of each 3D object
                        foreach (Model3DReferenceRelationshipPart part in embeddings_3d)
                        {
                            embedobj_number++;
                            Console.WriteLine($"--> Embedded object #{embedobj_number}");
                            Console.WriteLine($"----> Content Type: 3D model object");
                            Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                        }
                    }
                }
                return count_embedobj;
            }

            // Check for hyperlinks
            public List<Lists.Hyperlinks> Hyperlinks(string filepath)
            {
                List<Lists.Hyperlinks> results = new List<Lists.Hyperlinks>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    List<HyperlinkRelationship> hyperlinks = spreadsheet
                        .GetAllParts()
                        .SelectMany(p => p.HyperlinkRelationships)
                        .ToList();

                    foreach (HyperlinkRelationship hyperlink in hyperlinks)
                    {
                        // Add to list
                        results.Add(new Lists.Hyperlinks() { URL = hyperlink.Uri.ToString(), Action = Lists.ActionChecked });
                    }
                }
                return results;
            }

            // Check for printer settings
            public List<Lists.PrinterSettings> PrinterSettings(string filepath)
            {
                List<Lists.PrinterSettings> results = new List<Lists.PrinterSettings>();

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
                        results.Add(new Lists.PrinterSettings() { Uri = printer.Uri.ToString(), Action = Lists.ActionChecked });
                    }
                }
                return results;
            }

            // Check for active sheet
            public bool ActiveSheet(string filepath)
            {
                bool activeSheet = false;

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
                                if (workbookView.ActiveTab.Value > 0)
                                {
                                    activeSheet = true;
                                }
                            }
                        }
                    }
                }
                return activeSheet;
            }

            // Check for absolute path
            public List<Lists.AbsolutePath> AbsolutePath(string filepath)
            {
                List<Lists.AbsolutePath> results = new List<Lists.AbsolutePath>();
                AbsolutePath absPath = null;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    if (spreadsheet.WorkbookPart.Workbook.AbsolutePath != null)
                    {
                        absPath = spreadsheet.WorkbookPart.Workbook.GetFirstChild<AbsolutePath>();
                    }
                    // Add to list
                    results.Add(new Lists.AbsolutePath() { Path = absPath.ToString(), Action = Lists.ActionChecked });
                }
                return results;
            }

            // Check for metadata in file properties
            public List<Lists.FilePropertyInformation> FilePropertyInformation(string filepath)
            {
                List<Lists.FilePropertyInformation> results = new List<Lists.FilePropertyInformation>();
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
                    results.Add(new Lists.FilePropertyInformation() { Author = creator, Title = title, Subject = subject, Description = description, Keywords = keywords, Category = category, LastModifiedBy = lastmodifiedby, Found = found, Action = Lists.ActionChecked });
                }
                return results;
            }
        }
    }
}
