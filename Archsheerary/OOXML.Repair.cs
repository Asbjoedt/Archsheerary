using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Archsheerary
{
    public partial class OOXML
    {
        /// <summary>
        /// Collection of methods for reapairing Office Open XML spreadsheets.
        /// </summary>
        public class Repair
        {
            /// <summary>
            /// Repair OOXML spreadsheets with all repairs.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>True if file was repaired</returns>
            public static bool AllRepairs(string filepath)
            {
                bool success = false;

                bool repair_1 = Repair_VBA(filepath);
                bool repair_2 = Repair_DefinedNames(filepath);

                // If any repair method has been performed
                if (repair_1 == true && repair_2 == true)
                {
                    success = true;
                    return success;
                }
                return success;
            }

            /// <summary>
            /// Repair spreadsheets that had VBA code (macros) in them.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>True if file was repaired</returns>
            public static bool Repair_VBA(string filepath)
            {
                bool repaired = false;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    // Remove VBA project (if present) due to error in Open XML SDK
                    VbaProjectPart vba = spreadsheet.WorkbookPart.VbaProjectPart;
                    if (vba != null)
                    {
                        spreadsheet.WorkbookPart.DeletePart(vba);
                        repaired = true;
                    }
                }
                return repaired;
            }

            /// <summary>
            /// Repair invalid defined names. Returns true boolean if repair was succesful.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>True if file was repaired</returns>
            public static bool Repair_DefinedNames(string filepath)
            {
                bool repaired = false;

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    DefinedNames definedNames = spreadsheet.WorkbookPart.Workbook.DefinedNames;

                    // Remove legacy Excel 4.0 GET.CELL function (if present)
                    if (definedNames != null)
                    {
                        var definedNamesList = definedNames.ToList();
                        foreach (DefinedName definedName in definedNamesList)
                        {
                            if (definedName.InnerXml.Contains("GET.CELL"))
                            {
                                definedName.Remove();
                                repaired = true;
                            }
                        }
                    }

                    // Remove defined names with these " " (3 characters) in reference
                    if (definedNames != null)
                    {
                        var definedNamesList = definedNames.ToList();
                        foreach (DefinedName definedName in definedNamesList)
                        {
                            if (definedName.InnerXml.Contains("\" \""))
                            {
                                definedName.Remove();
                                repaired = true;
                            }
                        }
                    }
                }
                return repaired;
            }
        }
    }
}
