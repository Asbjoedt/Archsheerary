using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Archsheerary.Lists;

namespace Archsheerary
{
    public partial class Other
    {
        public class Enumerate
        {
            public static List<OriginalFilesIndex> Folder(string inputdir, bool recurse)
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
                List<OriginalFilesIndex> OriginalFilesList = new List<OriginalFilesIndex>();

                // Enrich metadata of each file and add to index of files if spreadsheet
                foreach (var entry in org_enumeration)
                {
                    FileInfo file_info = new FileInfo(entry);
                    if (FileFormatsIndex.Extension_Array.Contains(file_info.Extension) || FileFormatsIndex.Extension_Upper_Array.Contains(file_info.Extension))
                    {
                        string extension = file_info.Extension.ToLower();
                        string filename = file_info.Name;
                        string filepath = file_info.FullName;
                        OriginalFilesList.Add(new OriginalFilesIndex() { OriginalFilepath = filepath, OriginalFilename = filename, OriginalExtension = extension });
                    }
                }
                return OriginalFilesList;
            }
        }
    }
}
