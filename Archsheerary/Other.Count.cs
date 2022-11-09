using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Archsheerary
{
    public partial class Other
    {
        public class Count
        {
            // Count spreadsheets
            public List<Count> Spreadsheets(string input_dir, string output_dir, bool recurse)
            {
                //Object reference
                DirectoryInfo count = new DirectoryInfo(input_dir);
                Policy.FileFormats policyfileformats = new Policy.FileFormats();
                List<Lists.FileFormatsIndex> fileformats = policyfileformats.ListofFileFormats();
                List<Count> results = new List<Count>();

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                foreach (Lists.FileFormatsIndex fileformat in fileformats)
                {
                    // Count
                    int total = count.GetFiles($"*{fileformat.Extension}", searchoption).Length;

                    // Detect OOXML conformance
                    if (fileformat.Extension == ".xlsx")
                    {
                        if (fileformat.Conformance == "transitional")
                        {
                            total = CountOOXMLConformance(input_dir, recurse, fileformat.Conformance);
                        }
                        else if (fileformat.Conformance == "strict")
                        {
                            total = CountOOXMLConformance(input_dir, recurse, fileformat.Conformance);
                        }
                    }

                    // Change value in list
                    fileformat.Count = total;

                    // Create sum of all counts
                    numTOTAL = numTOTAL + total;

                    // Subtract if OOXML conformance was counted
                    if (fileformat.Conformance == "transitional" || fileformat.Conformance == "strict")
                    {
                        numTOTAL = numTOTAL - total;
                    }
                }

                // Inform user if no spreadsheets identified
                if (numTOTAL == 0)
                {
                    throw new Exception();
                }
                else
                {
                    return results;
                }
            }

            // Count XLSX Strict conformance
            public Tuple<int, int, int> CountOOXMLConformance(string input_directory, bool recurse)
            {
                string[] xlsx_files = { "" };
                int count_transitional = 0;
                int count_strict = 0;
                int check_fail = 0;

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                // Create index of xlsx files
                xlsx_files = Directory.GetFiles(input_directory, "*.xlsx", searchoption);
                foreach (var xlsx in xlsx_files)
                {
                    // Open each spreadsheet to check for conformance
                    try
                    {
                        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false))
                        {
                            Workbook workbook = spreadsheet.WorkbookPart.Workbook;

                            // Count Transitional
                            if (workbook.Conformance == null || workbook.Conformance == "transitional")
                            {
                                count_transitional++;
                            }
                            // Count Strict
                            else if (spreadsheet.StrictRelationshipFound == true)
                            {
                                count_strict++;
                            }
                        }

                    }
                    // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption
                    catch (InvalidDataException)
                    {
                        check_fail++;
                    }
                    catch (OpenXmlPackageException)
                    {
                        check_fail++;
                    }
                }

                // Return count as tuple
                return new Tuple<int, int, int>(count_transitional, count_strict, check_fail); 
            }
        }
    }
}
