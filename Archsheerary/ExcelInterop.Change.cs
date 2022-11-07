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
            // Change conformance to Strict
            public bool XLSXConformanceToStrict(string filepath)
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

                // If CLISC is run on Windows close Excel in task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task
                    Marshal.ReleaseComObject(app); // Delete Excel task
                }

                success = true;
                return success;
            }

            public bool XLSXConformanceToTransitional(string filepath)
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

                // If CLISC is run on Windows close Excel in task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task
                    Marshal.ReleaseComObject(app); // Delete Excel task
                }

                success = true;
                return success;
            }

            // Make first sheet active
            public List<Lists.ActiveSheet> ActivateFirstSheet(string filepath)
            {
                List<Lists.ActiveSheet> results = new List<Lists.ActiveSheet>();

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                try
                {
                    // Make first sheet active
                    if (app.Sheets.Count > 0)
                    {
                        Excel.Worksheet firstSheet = (Excel.Worksheet)app.ActiveWorkbook.Sheets[1];
                        firstSheet.Activate();
                        firstSheet.Select();

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
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Do nothing
                }
            }

        }
    }
}
