using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    /// <summary>
    /// Collection of other metods related to archiving spreadsheets
    /// </summary>
    public partial class Other
    {
        /// <summary>
        /// Collection of methods for checking validity of spreadsheets
        /// </summary>
        public class Check
        {
            /// <summary>
            /// Check for accepted file format extension
            /// </summary>
            public static bool Extension(string filepath)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (extension == ".xlsx" || extension == ".ods")
                {
                    validextension = true;
                }
                return validextension;
            }
        }
    }
}
