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
            /// <summary>
            /// Count number of spreadsheets in a folder with optional recurse parameter
            /// </summary>
            public List<Count> Spreadsheets(string input_dir, bool recurse)
            {
                DirectoryInfo count = new DirectoryInfo(input_dir);
                Other.FileFormats policy = new Other.FileFormats();
                List<DataTypes.FileFormatsIndex> fileformats = policy.FileFormatsIndex();
                List<Count> results = new List<Count>();

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                foreach (DataTypes.FileFormatsIndex fileformat in fileformats)
                {
                    // Count
                    int totalformat = count.GetFiles($"*{fileformat.Extension}", searchoption).Length;

                    // Change value in list
                    fileformat.Count = totalformat;

                    // Detect OOXML conformance
                    if (fileformat.Extension == ".xlsx")
                    {
                        Tuple<int, int, int> countedconformance = CountOOXMLConformance(input_dir, recurse);

                        if (fileformat.Conformance == "transitional")
                        {
                            fileformat.Count = countedconformance.Item1;
                        }

                        else if (fileformat.Conformance == "strict")
                        {
                            fileformat.Count = countedconformance.Item2;
                        }
                    }
                }
                return results;
            }

            /// <summary>
            /// Count XLSX spreadsheets with Strict conformance
            /// </summary>
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
