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
        /// Collection of methods for changing properties.
        /// </summary>
        public class Change
        {
            /// <summary>
            /// Set the file attributes to normal to change any write protection.
            /// </summary>
            public void FileAttributesProtection(string filepath)
            {
                File.SetAttributes(filepath, FileAttributes.Normal);
            }
        }
    }
}
