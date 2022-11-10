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
        public class Remove
        {
            /// <summary>
            /// Remove data connections using Excel interop
            /// </summary>
            public List<DataTypes.DataConnections> DataConnections(string filepath)
            {
                List<DataTypes.DataConnections> results = new List<DataTypes.DataConnections>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Remove data connections
                int count_conn = wb.Connections.Count;
                if (count_conn > 0)
                {
                    for (int i = 1; i <= wb.Connections.Count; i++)
                    {
                        // Add to list
                        var conn = wb.Connections[i];
                        results.Add(new DataTypes.DataConnections() { Name = conn.Name, Description = conn.Description, Type = conn.Type.ToString(), Action = DataTypes.ActionRemoved });

                        // Remove
                        wb.Connections[i].Delete();
                        i = i - 1;
                    }
                    count_conn = wb.Connections.Count;
                }

                // Save workbook and close Excel
                wb.Save();
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
            /// Remove external cell references using Excel interop
            /// </summary>
            public List<DataTypes.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<DataTypes.ExternalCellReferences> results = new List<DataTypes.ExternalCellReferences>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Find and replace external cell chains with cell values
                bool hasChain = false;
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
                                results.Add(new DataTypes.ExternalCellReferences() { Sheet = sheet.Name, Cell = cell.Address, Value = cell.Value.ToString(), Formula = cell.Formula.ToString(), Action = DataTypes.ActionRemoved });

                                hasChain = true;

                                // Remove
                                cell.Formula = "";
                                cell.Value2 = value;
                            }
                        }
                        if (hasChain == true)
                        {
                            // Save workbook and close Excel
                            wb.Save();
                            wb.Close();
                            app.Quit();

                            // If run on Windows release Excel from task manager
                            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                            {
                                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                            }
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
            /// Remove RealTimeData (RTD) functions using Excel interop
            /// </summary>
            public List<DataTypes.RTDFunctions> RTDFunctions(string filepath)
            {
                List<DataTypes.RTDFunctions> results = new List<DataTypes.RTDFunctions>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Find and replace RTD functions with cell values
                bool hasRTD = false;
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
                                results.Add(new DataTypes.RTDFunctions() { Sheet = sheet.Name, Cell = cell.Address, Value = cell.Value.ToString(), Formula = cell.Formula.ToString(), Action = DataTypes.ActionRemoved });

                                hasRTD = true;

                                // Remove
                                cell.Formula = "";
                                cell.Value2 = value;
                            }
                        }
                        if (hasRTD == true)
                        {
                            // Save workbook and close Excel
                            wb.Save();
                            wb.Close();
                            app.Quit();

                            // If run on Windows release Excel from task manager
                            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                            {
                                Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                                Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                            }
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
            /// Remove file property information using Excel interop
            /// </summary>
            public List<DataTypes.FilePropertyInformation> FilePropertyInformation(string filepath)
            {
                List<DataTypes.FilePropertyInformation> results = new List<DataTypes.FilePropertyInformation>();
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

                // Remove metadata
                if (wb.Author != null)
                {
                    creator = wb.Author;
                    wb.Author = "";
                }
                if (wb.Title != null)
                {
                    title = wb.Title;
                    wb.Title = "";
                }
                if (wb.Subject != null)
                {
                    subject = wb.Subject;
                    wb.Subject = "";
                }
                if (wb.Keywords != null)
                {
                    keywords = wb.Keywords;
                    wb.Keywords = "";
                }

                // Save workbook and close Excel
                wb.Save();
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Add to list
                results.Add(new DataTypes.FilePropertyInformation() { Author = creator, Title = title, Subject = subject, Description = description, Keywords = keywords, Category = category, LastModifiedBy = lastmodifiedby, Action = DataTypes.ActionRemoved });

                return results;
            }
        }
    }
}
