using System.IO;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Archsheerary
{
    public partial class ExcelInterop
    {
        public class Convert
        {
            // File format codes: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
            /// <summary>
            /// Convert any spreadsheet file format to another spreadsheet file format using Excel Interop
            /// </summary>
            public bool ToAnyFileFormat(string input_filepath, string output_filepath, int output_fileformat)
            {
                bool success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Save workbook and close Excel
                wb.SaveAs(output_filepath, output_fileformat);
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Repair spreadsheet
                OOXML.Repair rep = new OOXML.Repair();
                rep.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to XLSX Strict conformance using Excel Interop
            /// </summary>
            public bool ToXLSXStrict(string input_filepath, string output_filepath)
            {
                bool success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Save workbook and close Excel
                wb.SaveAs(output_filepath, 61);
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Repair spreadsheet
                OOXML.Repair rep = new OOXML.Repair();
                rep.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to XLSX Transitional conformance using Excel Interop
            /// </summary>
            public bool ToXLSXTransitional(string input_filepath, string output_filepath)
            {
                bool success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Save workbook and close Excel
                wb.SaveAs(output_filepath, 51);
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Repair spreadsheet
                OOXML.Repair rep = new OOXML.Repair();
                rep.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to ODS using Excel Interop
            /// </summary>
            public bool ToODS(string input_filepath, string output_filepath)
            {
                bool success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Save workbook and close Excel
                wb.SaveAs(output_filepath, 60);
                wb.Close();
                app.Quit();

                // If run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Repair spreadsheet
                OOXML.Repair rep = new OOXML.Repair();
                rep.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }
        }
    }
}
