using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Archsheerary
{
    public partial class OOXML
    {
        /// <summary>
        /// Collection of methods for converting Office Open XML spreadsheets
        /// </summary>
        public class Convert
        {
            /// <summary>
            /// Convert to XLSX Transitional conformance
            /// </summary>
            public static bool ToXLSXTransitional(string input_filepath, string output_filepath)
            {
                bool convert_success = false;


                // If protected in file properties
                File.SetAttributes(input_filepath, FileAttributes.Normal); // Remove file attributes on spreadsheet
                // If password-protected or reserved by another user
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(input_filepath, false))
                {
                    if (spreadsheet.WorkbookPart.Workbook.WorkbookProtection != null || spreadsheet.WorkbookPart.Workbook.FileSharing != null)
                    {
                        return convert_success;
                    }
                }

                // Convert spreadsheet
                byte[] byteArray = File.ReadAllBytes(input_filepath);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                    {
                        spreadsheet.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                    }
                    File.WriteAllBytes(output_filepath, stream.ToArray());
                }

                // Repair spreadsheet
                OOXML.Repair.AllRepairs(output_filepath);

                // Return success
                convert_success = true;
                return convert_success;
            }
        }
    }
}
