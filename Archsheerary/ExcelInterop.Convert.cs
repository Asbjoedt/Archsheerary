﻿using System.IO;
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
            /// Convert any spreadsheet file format to another spreadsheet file format using Excel Interop.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_filepath">Path to output file</param>
            /// <param name="output_extension">Extension of the output file format</param>
            /// <returns>True if conversion was successful</returns>
            /// <exception cref="Exception">Thrown if file format extension was not recognized.</exception>
            public static bool ToAnySpreadsheetFileFormat(string input_filepath, string output_filepath, string output_extension)
            {
                bool success = false;
                int? excel_int = null;

                // Get real output file format
                List<DataTypes.FileFormatsIndex> fileFormats = Other.FileFormats.FileFormatsIndex();
                foreach (var fileFormat in fileFormats)
                {
                    if (output_extension.ToLower() == fileFormat.Extension && fileFormat.Conformance == null)
                    {
                        excel_int = fileFormat.ExcelFileFormat;
                    }
                    else if (fileFormat.Conformance == "transitional")
                    {
                        excel_int = 51;
                    }
                    else if (fileFormat.Conformance == "strict")
                    {
                        excel_int = 61;
                    }
                }
                if (excel_int == null)
                {
                    throw new Exception("Output file format extension was not recognized");
                }

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Save workbook and close Excel
                wb.SaveAs(output_filepath, excel_int);
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
            /// Convert to XLSX Strict conformance using Excel Interop.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_filepath">Path to output file</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToXLSXStrict(string input_filepath, string output_filepath)
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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to XLSX Transitional conformance using Excel Interop.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_filepath">Path to output file</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToXLSXTransitional(string input_filepath, string output_filepath)
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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }

            /// <summary>
            /// Convert to ODS using Excel Interop.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_filepath">Path to output file</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToODS(string input_filepath, string output_filepath)
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
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                success = true;
                return success;
            }
        }
    }
}
