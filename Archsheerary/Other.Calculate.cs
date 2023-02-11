using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;

namespace Archsheerary
{
    public partial class Other
    {
        /// <summary>
        /// Collection of methods for creating checksums of files.
        /// </summary>
        public class Calculate
        {
            /// <summary>
            /// Calculate MD5 checksum of file.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <returns>String with MD5 checksum</returns>
            public static string MD5Hash(string filepath)
            {
                using (var md5 = MD5.Create())
                {
                    using (var stream = File.OpenRead(filepath))
                    {
                        var checksum = md5.ComputeHash(stream);
                        return BitConverter.ToString(checksum).Replace("-", "").ToLowerInvariant();
                    }
                }
            }
        }
    }
}
