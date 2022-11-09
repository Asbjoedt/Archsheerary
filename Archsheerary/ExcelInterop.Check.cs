using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Archsheerary
{
    public partial class ExcelInterop
    {
        public class Check
        {
            // Check for data connections
            public List<Lists.DataConnections> DataConnections(string filepath)
            {
                List<Lists.DataConnections> results = new List<Lists.DataConnections>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                int count_conn = wb.Connections.Count;
                if (count_conn > 0)
                {
                    for (int i = 1; i <= wb.Connections.Count; i++)
                    {
                        results.Add(new Lists.DataConnections() { Description = wb.Connections., Action = Lists.ActionChecked });
                    }
                }

                // Close Excel
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }
                return results;
            }

            // Check for external cell references
            public List<Lists.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<Lists.ExternalCellReferences> results = new List<Lists.ExternalCellReferences>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                foreach (Excel.Worksheet sheet in wb.Sheets)
                {
                    try
                    {
                        Excel.Range range = (Excel.Range)sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                        foreach (Excel.Range cell in range.Cells)
                        {
                            var value = cell.Value2;
                            string formula = cell.Formula.ToString();
                            string hit = formula.Substring(0, 2); // Transfer first 2 characters to string

                            if (hit == "='")
                            {
                                // Add to list
                                results.Add(new Lists.ExternalCellReferences() { Sheet = sheet.Name, Cell = cell.Address, Value = cell.Value.ToString(), Formula = cell.Formula.ToString(), Action = Lists.ActionChecked });
                            }
                        }

                        // Close Excel
                        wb.Close();
                        app.Quit();

                        // If run on Windows release Excel from task manager
                        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                        {
                            Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                            Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException) // Catch if no formulas in range
                    {
                        // Do nothing
                    }
                    catch (System.ArgumentOutOfRangeException) // Catch if formula has less than 2 characters
                    {
                        // Do nothing
                    }

                }
                return results;
            }

            // Check for RTD functions
            public List<Lists.RTDFunctions> RTDFunctions(string filepath)
            {
                List<Lists.RTDFunctions> results = new List<Lists.RTDFunctions>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                foreach (Excel.Worksheet sheet in wb.Sheets)
                {
                    try
                    {
                        Excel.Range range = (Excel.Range)sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                        foreach (Excel.Range cell in range.Cells)
                        {
                            var value = cell.Value2;
                            string formula = cell.Formula.ToString();
                            string hit = formula.Substring(0, 4); // Transfer first 4 characters to string
                            if (hit == "=RTD")
                            {
                                // Add to list
                                results.Add(new Lists.RTDFunctions() { Sheet = sheet.Name, Cell = cell.Address, Value = cell.Value.ToString(), Formula = cell.Formula.ToString(), Action = Lists.ActionChecked });
                            }
                        }

                        // Close Excel
                        wb.Close();
                        app.Quit();

                        // If run on Windows release Excel from task manager
                        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                        {
                            Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                            Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException) // Catch if no formulas in range
                    {
                        // Do nothing
                    }
                }
                return results;
            }

            // CHeck for active sheet
            public List<Lists.ActiveSheet> ActiveSheet(string filepath)
            {
                List<Lists.ActiveSheet> results = new List<Lists.ActiveSheet>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                if (app.ActiveSheet != app.ActiveWorkbook.Sheets[1])
                {
                    // Add to list
                    results.Add(new Lists.ActiveSheet() { ActiveSheeet = (uint)app.ActiveSheet, Action = Lists.ActionChecked });

                    // Close Excel
                    wb.Close();
                    app.Quit();

                    // If run on Windows release Excel from task manager
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                        Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                    }
                }
                return results;
            }

            // Check for file property information
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

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                if (wb.Author != null)
                {
                    creator = wb.Author;
                }
                if (wb.Title != null)
                {
                    title = wb.Title;
                }
                if (wb.Subject != null)
                {
                    subject = wb.Subject;
                }
                if (wb.Keywords != null)
                {
                    keywords = wb.Keywords;
                }

                // Close Excel
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Add to list
                results.Add(new Lists.FilePropertyInformation() { Author = creator, Title = title, Subject = subject, Description = description, Keywords = keywords, Category = category, LastModifiedBy = lastmodifiedby, Action = Lists.ActionChecked });

                return results;
            }
        }
    }
}
