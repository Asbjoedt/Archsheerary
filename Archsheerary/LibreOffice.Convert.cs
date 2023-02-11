using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    /// <summary>
    /// Collection of methods using LibreOffice.
    /// </summary>
    public class LibreOffice
    {
        /// <summary>
        /// Collection of methods for converting spreadsheets.
        /// </summary>
        public class Convert
        {
            /// <summary>
            /// Convert spreadsheets to any other spreadsheet file format using LibreOffice.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <param name="output_extension">Extension of the output file format</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToAnySpreadsheetFileFormat(string input_filepath, string output_folder, string output_extension)
            {
                bool success = false;
                Process app = new Process();

                // If output extension begins with a dot, then remove dot
                output_extension = output_extension.ToLower().Split(".").Last();

                // If app is run on Windows
                string? dir = null;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    dir = Environment.GetEnvironmentVariable("LibreOffice");
                }
                if (dir != null)
                {
                    app.StartInfo.FileName = dir;
                }
                else
                {
                    app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
                }

                app.StartInfo.Arguments = $"--headless --convert-to {output_extension}  {input_filepath} --outdir {output_folder}";
                app.Start();
                app.WaitForExit();
                app.Close();

                success = true;
                return success;
            }

            /// <summary>
            /// Convert spreadsheets to ODS file format using LibreOffice. Returns true boolean if successful conversion.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToODS(string input_filepath, string output_folder)
            {
                bool success = false;
                Process app = new Process();

                // If app is run on Windows
                string? dir = null;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    dir = Environment.GetEnvironmentVariable("LibreOffice");
                }
                if (dir != null)
                {
                    app.StartInfo.FileName = dir;
                }
                else
                {
                    app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
                }

                app.StartInfo.Arguments = "--headless --convert-to ods " + input_filepath + " --outdir " + output_folder;
                app.Start();
                app.WaitForExit();
                app.Close();

                success = true;
                return success;
            }

            /// <summary>
            /// Convert spreadsheets to XLSX Transitional conformance file format using LibreOffice
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_folder">Path to output folder</param>
            /// <returns>True if conversion was successful</returns>
            public static bool ToXLSXTransitional(string input_filepath, string output_folder)
            {
                bool success = false;
                Process app = new Process();

                // If app is run on Windows
                string? dir = null;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    dir = Environment.GetEnvironmentVariable("LibreOffice");
                }
                if (dir != null)
                {
                    app.StartInfo.FileName = dir;
                }
                else
                {
                    app.StartInfo.FileName = "C:\\Program Files\\LibreOffice\\program\\scalc.exe";
                }

                app.StartInfo.Arguments = "--headless --convert-to xlsx " + input_filepath + " --outdir " + output_folder;
                app.Start();
                app.WaitForExit();
                app.Close();

                success = true;
                return success;
            }
        }
    }
}