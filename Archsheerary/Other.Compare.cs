using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public partial class Other
    {
        public class Compare
        {
            // Use Beyond Compare 4 command line for comparison
            public bool CompareWorkbook(string filepath_one, string filepath_two)
            {
                Process app = new Process();
                bool success = false;

                // If run on Windows
                string? dir = null;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
                {
                    dir = Environment.GetEnvironmentVariable("BeyondCompare");
                }
                if (dir != null)
                {
                    app.StartInfo.FileName = dir;
                }
                else
                {
                    app.StartInfo.FileName = "C:\\Program Files\\Beyond Compare 4\\BCompare.exe";
                }

                // Run program
                app.StartInfo.Arguments = $"\"{filepath_one}\" \"{filepath_two}\" /silent /qc=<crc> /ro";
                app.Start();
                app.WaitForExit();
                int returncode = app.ExitCode;
                app.Close();

                // Handle ExitCode
                if (returncode == 0 || returncode == 1 || returncode == 2)
                {
                    success = true;
                }
                if (returncode == 12 || returncode == 13 || returncode == 14)
                {
                    success = false;
                }
                if (returncode == 11 || returncode == 100 || returncode == 104)
                {
                    throw new Exception("Error in comparison of two spreadsheets occured");
                }
                return success;
            }
        }
    }
}
