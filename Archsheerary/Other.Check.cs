using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public partial class Other
    {
        public class Check
        {
            /// <summary>
            /// Check for accepted file format extension
            /// </summary>
            public bool Extension(string filepath)
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
