using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public partial class Other
    {
        public class Enumerate
        {
            /// <summary>
            /// Enumerate all files in a folder with optional recurse parameter
            /// </summary>
            public static List<Lists.OriginalFilesIndex> Folder(string inputdir, bool recurse)
            {
                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                // Enumerate input directory
                IEnumerable<string> org_enumeration = Directory.EnumerateFiles(inputdir, "*", searchoption).ToList();

                // Create new fileIndex for spreadsheets
                List<Lists.OriginalFilesIndex> OriginalFilesList = new List<Lists.OriginalFilesIndex>();

                // Create list and subsequently array of spreadsheet file formats
                Policy.FileFormats policy = new Policy.FileFormats();
                List<Lists.FileFormatsIndex> FileFormats = policy.ListofFileFormats();
                string[] fileformats = FileFormats.ToArray();

                // Enrich metadata of each file and add to index of files if spreadsheet
                foreach (var entry in org_enumeration)
                {
                    FileInfo file_info = new FileInfo(entry);
                    if (FileFormats.Extension.Contains(file_info.Extension) || FileFormats.ExtensionUpper.Contains(file_info.Extension))
                    {
                        string extension = file_info.Extension.ToLower();
                        string filename = file_info.Name;
                        string filepath = file_info.FullName;
                        OriginalFilesList.Add(new Lists.OriginalFilesIndex() { OriginalFilepath = filepath, OriginalFilename = filename, OriginalExtension = extension });
                    }
                }
                return OriginalFilesList;
            }
        }
    }
}
