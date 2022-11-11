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
        /// <summary>
        /// Collection of methods for enumerating data
        /// </summary>
        public class Enumerate
        {
            /// <summary>
            /// Enumerate all files in a folder with optional recurse parameter
            /// </summary>
            public static List<DataTypes.OriginalFilesIndex> Folder(string input_directory, bool recurse)
            {
                // Search recursively or not
                SearchOption searchoption = SearchOption.TopDirectoryOnly;
                if (recurse == true)
                {
                    searchoption = SearchOption.AllDirectories;
                }

                // Enumerate input directory
                IEnumerable<string> OriginalFilesEnumeration = Directory.EnumerateFiles(input_directory, "*", searchoption).ToList();

                // Create new fileIndex for spreadsheets
                List<DataTypes.OriginalFilesIndex> OriginalFilesList = new List<DataTypes.OriginalFilesIndex>();

                // Create list and subsequently an array of spreadsheet file formats
                List<DataTypes.FileFormatsIndex> FileFormats = Other.FileFormats.FileFormatsIndex();
                string[] extensionArray = {""};
                string[] extensionUpperArray = {""};
                foreach (DataTypes.FileFormatsIndex fileformat in FileFormats)
                {
                    extensionArray =  fileformat.Extension.ToCharArray().Select(c => c.ToString()).ToArray();
                    extensionUpperArray = fileformat.Extension.ToCharArray().Select(c => c.ToString()).ToArray();
                }

                // Enrich metadata of each file and add to index of files if spreadsheet
                foreach (var file in OriginalFilesEnumeration)
                {
                    FileInfo fileinfo = new FileInfo(file);
                    if (extensionArray.Contains(fileinfo.Extension) || extensionUpperArray.Contains(fileinfo.Extension))
                    {
                        OriginalFilesList.Add(new DataTypes.OriginalFilesIndex() { OriginalFilepath = fileinfo.FullName, OriginalFilename = fileinfo.Name, OriginalExtension = fileinfo.Extension, OriginalExtensionLower = fileinfo.Extension.ToLower() });
                    }
                }
                return OriginalFilesList;
            }
        }
    }
}
