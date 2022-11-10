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
        public class Checksum
        {
            /// <summary>
            /// Calculate MD5 checksum of file
            /// </summary>
            public string MD5Hash(string filepath)
            {
                try
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
                // If no converted spreadsheet exist
                catch (System.ArgumentException)
                {
                    return "";
                }
                catch (System.IO.FileNotFoundException)
                {
                    return "";
                }
            }
        }
    }
}
