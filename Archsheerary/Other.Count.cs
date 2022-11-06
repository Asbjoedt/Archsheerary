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
        public partial class Count
        {
            // Count spreadsheets
            public List<Count> Spreadsheets(string input_dir, string output_dir, bool recurse)
            {
                //Object reference
                DirectoryInfo count = new DirectoryInfo(input_dir);
                List<Lists.FileFormatsIndex> fileformats = new List<Lists.FileFormatsIndex>();
                List<Count> results = new List<Count>();

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                foreach (FileFormatsIndex fileformat in fileformats)
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
            public static int numCONFORM_fail = 0;

            // Count XLSX Strict conformance
            public int CountOOXMLConformance(string inputdir, bool recurse, string conformance)
            {
                int count = 0;
                string[] xlsx_files = { "" };

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                // Create index of xlsx files
                xlsx_files = Directory.GetFiles(inputdir, "*.xlsx", searchoption);

                // Open each spreadsheet to check for conformance
                try
                {
                    // Count Transitional
                    if (conformance == "transitional")
                    {
                        foreach (var xlsx in xlsx_files)
                        {
                            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false))
                            {
                                Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                                if (workbook.Conformance == null || workbook.Conformance == "transitional")
                                {
                                    count++;
                                }
                            }
                        }
                    }
                    // Count Strict
                    else if (conformance == "strict")
                    {
                        foreach (var xlsx in xlsx_files)
                        {
                            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false);
                            bool? strict = spreadsheet.StrictRelationshipFound;
                            spreadsheet.Close();
                            if (strict == true)
                            {
                                count++;
                            }
                        }
                    }

                }

                // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption
                catch (InvalidDataException)
                {
                    numCONFORM_fail++;
                }
                catch (OpenXmlPackageException)
                {
                    numCONFORM_fail++;
                }

                // Return count
                return count;
            }
        }
    }
}
