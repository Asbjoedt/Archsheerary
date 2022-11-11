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
        /// Copy spreadsheet
        /// </summary>
        public class Copy
        {
            /// <summary>
            /// Copy file to another location (output filepath must include full folder path, filename and file extension)
            /// </summary>
            public static void Spreadsheet(string input_filepath, string output_filepath)
            {
                File.Copy(input_filepath, output_filepath);
                File.SetAttributes(output_filepath, FileAttributes.Normal); // Remove file attributes on copied file
            }
        }
    }
}
