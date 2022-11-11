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
        public class Change
        {
            /// <summary>
            /// Change conformance of XLSX file to Strict
            /// </summary>
            public static bool XLSXConformanceToStrict(string filepath)
            {
                bool success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Convert to Strict and close Excel
                wb.SaveAs(filepath, 61);
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task
                    Marshal.ReleaseComObject(app); // Delete Excel task
                }

                success = true;
                return success;
            }

            /// <summary>
            /// Change conformance of XLSX file to Transitional
            /// </summary>
            public static bool XLSXConformanceToTransitional(string filepath)
            {
                bool success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Convert to Strict and close Excel
                wb.SaveAs(filepath, 51);
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task
                    Marshal.ReleaseComObject(app); // Delete Excel task
                }

                success = true;
                return success;
            }

            /// <summary>
            /// Make first sheet active
            /// </summary>
            public static List<DataTypes.ActiveSheet> ActivateFirstSheet(string filepath)
            {
                List<DataTypes.ActiveSheet> results = new List<DataTypes.ActiveSheet>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Make first sheet active
                if (app.ActiveSheet != app.ActiveWorkbook.Sheets[1])
                {
                    // Add to list
                    results.Add(new DataTypes.ActiveSheet() { OriginalActiveSheet = (uint)app.ActiveSheet, NewActiveSheet = 0, Action = DataTypes.ActionChanged });

                    // Change
                    Excel.Worksheet firstSheet = (Excel.Worksheet)app.ActiveWorkbook.Sheets[1];
                    firstSheet.Activate();
                    firstSheet.Select();

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
                return results;
            }
        }
    }
}
