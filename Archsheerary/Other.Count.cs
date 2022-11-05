using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;

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
                fileFormatIndex index = new fileFormatIndex();
                List<fileFormatIndex> fileformats = index.Create_fileFormatIndex();
                List<Count> results = new List<Count>();

                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                foreach (fileFormatIndex fileformat in fileformats)
                {
                    // Count
                    int total = count.GetFiles($"*{fileformat.Extension}", searchoption).Length;

                    // Detect OOXML conformance
                    if (fileformat.Extension == ".xlsx")
                    {
                        if (fileformat.Conformance == "transitional")
                        {
                            total = Count_OOXML_Conformance(input_dir, recurse, fileformat.Conformance);
                        }
                        else if (fileformat.Conformance == "strict")
                        {
                            total = Count_OOXML_Conformance(input_dir, recurse, fileformat.Conformance);
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
        }
    }
}
