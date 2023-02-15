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
        /// <summary>
        /// Collection of methods for checking content in spreadsheets.
        /// </summary>
        public class Check
        {
            /// <summary>
            /// Check for data connections.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified data connections</returns>
            public static List<DataTypes.DataConnections> DataConnections(string filepath)
            {
                List<DataTypes.DataConnections> results = new List<DataTypes.DataConnections>();

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
                        var conn = wb.Connections[i];
                        results.Add(new DataTypes.DataConnections() { Name = conn.Name, Description = conn.Description, Type = conn.Type.ToString(), Action = DataTypes.ActionChecked });
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

            /// <summary>
            /// Check for external cell references.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified external cell references</returns>
            /// <exception cref="COMException">Thrown if no formulas in range.</exception>
            /// <exception cref="ArgumentOutOfRangeException">Thrown if formula has less than 2 characters.</exception>
            public static List<DataTypes.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<DataTypes.ExternalCellReferences> results = new List<DataTypes.ExternalCellReferences>();

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
                                results.Add(new DataTypes.ExternalCellReferences() { Sheet = sheet.Name, Cell = cell.Address, Value = cell.Value.ToString(), Formula = cell.Formula.ToString(), Action = DataTypes.ActionChecked });
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

            /// <summary>
            /// Check for RealTimeData (RTD) functions.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified RTD functions</returns>
            /// <exception cref="COMException">Thrown if no formulas in range.</exception>
            public static List<DataTypes.RTDFunctions> RTDFunctions(string filepath)
            {
                List<DataTypes.RTDFunctions> results = new List<DataTypes.RTDFunctions>();

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
                                results.Add(new DataTypes.RTDFunctions() { Sheet = sheet.Name, Cell = cell.Address, Value = cell.Value.ToString(), Formula = cell.Formula.ToString(), Action = DataTypes.ActionChecked });
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

            /// <summary>
            /// Check for active sheet.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified active sheet</returns>
            public static List<DataTypes.ActiveSheet> ActiveSheet(string filepath)
            {
                List<DataTypes.ActiveSheet> results = new List<DataTypes.ActiveSheet>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                if (app.ActiveSheet != app.ActiveWorkbook.Sheets[1])
                {
                    // Add to list
                    results.Add(new DataTypes.ActiveSheet() { OriginalActiveSheet = (uint)app.ActiveSheet, NewActiveSheet = null, Action = DataTypes.ActionChecked });

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

            /// <summary>
            /// Check for file property information.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>List of identified file property information</returns>
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

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Perform check
                if (wb.Author != null)
                {
                    creator = wb.Author;
                    found = true;
                }
                if (wb.Title != null)
                {
                    title = wb.Title;
                    found = true;
                }
                if (wb.Subject != null)
                {
                    subject = wb.Subject;
                    found = true;
                }
                if (wb.Keywords != null)
                {
                    keywords = wb.Keywords;
                    found = true;
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
                results.Add(new DataTypes.FilePropertyInformation() { Author = creator, Title = title, Subject = subject, Description = description, Keywords = keywords, Category = category, LastModifiedBy = lastmodifiedby, FilePropertyInfoFound = found, Action = DataTypes.ActionChecked });

                return results;
            }
        }
    }
}
