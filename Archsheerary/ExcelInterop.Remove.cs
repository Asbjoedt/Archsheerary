﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using static Archsheerary.Lists;

namespace Archsheerary
{
    public partial class ExcelInterop
    {
        public class Remove
        {
            // Remove data connections
            public List<Lists.DataConnections> DataConnections(string filepath)
            {
                List<Lists.DataConnections> results = new List<Lists.DataConnections>();

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
                        results.Add(new Lists.DataConnections() { Description = wb.Connections., Action = "Removed" });

                        wb.Connections[i].Delete();
                        i = i - 1;
                    }
                    count_conn = wb.Connections.Count;
                }

                // Save workbook and close Excel
                wb.Save();
                wb.Close();
                app.Quit();

                // If CLISC is run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }
                return results;
            }

            // Remove external cell references
            public List<Policy.ExternalCellReferences> ExternalCellReferences(string filepath)
            {
                List<Policy.ExternalCellReferences> results = new List<Policy.ExternalCellReferences>();

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
                                hasChain = true;
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

                            // If CLISC is run on Windows release Excel from task manager
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

            // Remove RTD functions
            public void RTDFunctions(string filepath)
            {
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
                                cell.Formula = "";
                                cell.Value2 = value;
                                hasRTD = true;
                            }
                        }
                        if (hasRTD = true)
                        {
                            // Save workbook and close Excel
                            wb.Save();
                            wb.Close();
                            app.Quit();

                            // If CLISC is run on Windows release Excel from task manager
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
            }

            public void Metadata(string filepath)
            {
                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Remove metadata
                if (wb.Author != null)
                {
                    wb.Author = "";
                }
                if (wb.Title != null)
                {
                    wb.Title = "";
                }
                if (wb.Subject != null)
                {
                    wb.Subject = "";
                }
                if (wb.Keywords != null)
                {
                    wb.Keywords = "";
                }

                // Save workbook and close Excel
                wb.Save();
                wb.Close();
                app.Quit();

                // If CLISC is run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }
            }
        }
    }
}
