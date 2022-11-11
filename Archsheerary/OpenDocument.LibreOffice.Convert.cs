using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public partial class OpenDocument
    {
        public class LibreOffice
        {
            public class Convert
            {
                /// <summary>
                /// Convert spreadsheets to any other spreadsheet file format using LibreOffice
                /// </summary>
                public static bool ToAnyFileFormat(string input_filepath, string output_folder, string output_fileformat)
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

                    app.StartInfo.Arguments = $"--headless --convert-to {output_fileformat}  {input_filepath} --outdir {output_folder}";
                    app.Start();
                    app.WaitForExit();
                    app.Close();

                    success = true;
                    return success;
                }

                /// <summary>
                /// Convert spreadsheets to ODS file format using LibreOffice
                /// </summary>
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
}