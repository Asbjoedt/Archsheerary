using System.IO;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Archsheerary
{
    public partial class ExcelInterop
    {
        /// <summary>
        /// Collection of methods for converting spreadsheets.
        /// </summary>
        public class Convert
        {
            /// <summary>
            /// Convert any spreadsheet file format to another spreadsheet file format using Excel Interop. Returns true boolean is conversion was succesful.
            /// Output file formats: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
            /// </summary>
            public static bool ToAnySpreadsheetFileFormat(string input_filepath, string output_filepath, int output_fileformat)
            {
                bool success = false;

                // If protected in file properties
                File.SetAttributes(input_filepath, FileAttributes.Normal); // Remove file attributes on spreadsheet

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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to XLSX Strict conformance using Excel Interop. Returns true boolean is conversion was succesful.
            /// </summary>
            public static bool ToXLSXStrict(string input_filepath, string output_filepath)
            {
                bool success = false;

                // If protected in file properties
                File.SetAttributes(input_filepath, FileAttributes.Normal); // Remove file attributes on spreadsheet

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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to XLSX Transitional conformance using Excel Interop. Returns true boolean is conversion was succesful.
            /// </summary>
            public static bool ToXLSXTransitional(string input_filepath, string output_filepath)
            {
                bool success = false;

                // If protected in file properties
                File.SetAttributes(input_filepath, FileAttributes.Normal); // Remove file attributes on spreadsheet

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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to ODS using Excel Interop. Returns true boolean is conversion was succesful.
            /// </summary>
            public static bool ToODS(string input_filepath, string output_filepath)
            {
                bool success = false;

                // If protected in file properties
                File.SetAttributes(input_filepath, FileAttributes.Normal); // Remove file attributes on spreadsheet

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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }
        }
    }
}
