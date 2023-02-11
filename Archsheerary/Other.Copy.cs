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
        /// Copy spreadsheet.
        /// </summary>
        public class Copy
        {
            /// <summary>
            /// Copy file to another location.
            /// </summary>
            /// <param name="input_filepath">Path to input file</param>
            /// <param name="output_filepath">Path to output file</param>
            public static void Spreadsheet(string input_filepath, string output_filepath)
            {
                File.Copy(input_filepath, output_filepath);
                File.SetAttributes(output_filepath, FileAttributes.Normal); // Remove file attributes on copied file
            }
        }
    }
}
