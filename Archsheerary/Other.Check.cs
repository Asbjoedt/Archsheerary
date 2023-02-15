using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using Excel = Microsoft.Office.Interop.Excel;

namespace Archsheerary
{
    /// <summary>
    /// Collection of other metods related to archiving spreadsheets.
    /// </summary>
    public partial class Other
    {
        /// <summary>
        /// Collection of methods for checking validity of spreadsheets.
        /// </summary>
        public class Check
        {
            /// <summary>
            /// Check for OOXML and OpenDocument file format extensions. Returns true boolean if valid extension.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="normalize">True if only .xlsx and .ods extensions are valid</param>
            /// <returns>True bool if extension is valid</returns>
            public static bool ExtensionOOXMLAndOpenDocument(string filepath, bool normalize)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (!normalize)
                {
                    if (extension == ".xlam" || extension == ".xlsx" || extension == ".xlsm" || extension == ".xltm" || extension == ".xlsx" || extension == ".fods" || extension == ".ods" || extension == ".ots")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                else if (normalize)
                {
                    if (extension == ".xlsx" || extension == ".ods")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                return validextension;
            }

            /// <summary>
            /// Check for OOXML extensions. Returns true boolean if valid extension.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="normalize">True if only .xlsx extension is valid</param>
            /// <return>True bool if extension is valid</return>
            public static bool ExtensionOOXML(string filepath, bool normalize)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (!normalize)
                {
                    if (extension == ".xlam" || extension == ".xlsx" || extension == ".xlsm" || extension == ".xltm" || extension == ".xlsx")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                else if (normalize)
                {
                    if (extension == ".xlsx")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                return validextension;
            }

            /// <summary>
            /// Check for OpenDocument Spreadsheets extensions. Returns true boolean if valid extension.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="normalize">True if only .ods extension is valid</param>
            /// <return>True bool if extension is valid</return>
            public static bool ExtensionOpenDocument(string filepath, bool normalize)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (!normalize)
                {
                    if (extension == ".fods" || extension == ".ods" || extension == ".ots")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                else if (normalize)
                {
                    if (extension == ".ods")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                return validextension;
            }

            /// <summary>
            /// Check if fileattributes protects the file from being written to.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <return>True bool if file is protected</return>
            public bool FileAttributesProtection(string filepath)
            {
                bool protect = false;

                // Get file attributes and check if read-only
                FileAttributes filattri = File.GetAttributes(filepath);
                if (filattri.HasFlag(FileAttributes.ReadOnly))
                {
                    protect = true;
                    return protect;
                }
                return protect;
            }

            /// <summary>
            /// Check if password protects the file from being written to by trying to open the file.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <return>True bool if file is password protected</return>
            /// <exception cref="FileFormatException">Thrown if an OOXML file is password protected.</exception>
            /// <exception cref="COMException">Thrown if OpenDocument or Numbers file is password protected.</exception>
            public bool PasswordProtection(string filepath)
            {
                bool protect = false;
                string extension = Path.GetExtension(filepath).ToLower();

                // Perform check by trying to open the file by an application
                try
                {
                    if (extension == ".xlam" || extension == ".xlsm" || extension == ".xlsx" || extension == ".xltm" || extension == ".xltx")
                    {
                        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                        {
                            // Do nothing
                        }
                    }
                    else if (extension == ".fods" || extension == ".ods" || extension == ".ots" || extension == ".numbers")
                    {
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

                        app.StartInfo.Arguments = "--calc " + filepath;
                        app.Start();
                        app.WaitForExit();
                        app.Close();
                    }
                    else if (extension == ".xla" || extension == ".xls" || extension == ".xlt" || extension == ".xlsb")
                    {
                        // Open Excel
                        Excel.Application app = new Excel.Application();
                        app.DisplayAlerts = false;
                        Excel.Workbook wb = app.Workbooks.Open(filepath, Notify: false);

                        // Close Excel
                        wb.Close();
                        app.Quit();
                    }
                }
                catch (System.IO.FileFormatException) // OOXML catch
                {
                    protect = true;
                    return protect;
                }
                catch (System.Runtime.InteropServices.COMException) // ExcelInterop catch
                {
                    protect = true;
                    return protect;
                }
                return protect;
            }
        }
    }
}
