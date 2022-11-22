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
        /// <summary>
        /// Collection of methods for counting spreadsheets.
        /// </summary>
        public class Count
        {
            /// <summary>
            /// Count number of spreadsheets in a folder with optional recurse parameter. Returns list of counted spreadsheet file formats.
            /// </summary>
            public static List<Count> Spreadsheets(string input_directory, bool recurse)
            {
                DirectoryInfo count = new DirectoryInfo(input_directory);
                List<DataTypes.FileFormatsIndex> fileformats = Other.FileFormats.FileFormatsIndex();
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
                        Tuple<int, int, int> countedconformance = OOXMLConformance(input_directory, recurse);

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
            /// Count XLSX spreadsheets based on conformance. Returns tuple with Transitional count, Strict count and failed count.
            /// </summary>
            public static Tuple<int, int, int> OOXMLConformance(string input_directory, bool recurse)
            {
                string[] xlsx_files = { "" };
                int count_transitional = 0;
                int count_strict = 0;
                int count_fail = 0;

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
                    // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption.
                    catch (InvalidDataException)
                    {
                        count_fail++;
                    }
                    catch (OpenXmlPackageException)
                    {
                        count_fail++;
                    }
                }

                // Return count as tuple
                return new Tuple<int, int, int>(count_transitional, count_strict, count_fail); 
            }

            /// <summary>
            /// Count XLSX spreadsheets with Strict conformance. Returns tuple with Strict count and failed count.
            /// </summary>
            public Tuple<int, int> StrictConformance(string input_directory, bool recurse)
            {
                int count_strict = 0;
                int count_fail = 0;
                string[] xlsx_files = { "" };

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                // Create index of xlsx files
                xlsx_files = Directory.GetFiles(input_directory, "*.xlsx", searchoption);

                try
                {
                    foreach (var xlsx in xlsx_files)
                    {
                        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(xlsx, false);
                        bool? strict = spreadsheet.StrictRelationshipFound;
                        spreadsheet.Close();
                        if (strict == true)
                        {
                            count_strict++;
                        }
                    }
                }
                // Catch exceptions, when spreadsheet cannot be opened due to password protection or corruption
                catch (InvalidDataException)
                {
                    count_fail++;
                }
                catch (OpenXmlPackageException)
                {
                    count_fail++;
                }
                // Return count as tuple
                return new Tuple<int, int>(count_strict, count_fail);
            }
        }
    }
}
