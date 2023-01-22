using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    /// <summary>
    /// Collection of methods for OpenDocument Spreadsheets.
    /// </summary>
    public partial class OpenDocument
    {
        /// <summary>
        /// Collection of methods for validating OpenDocument spreadsheets.
        /// </summary>
        public partial class Validate
        {
            /// <summary>
            /// Validate OpenDocument Spreadsheets using ODF Validator. Returns true boolean if valid.
            /// </summary>
            public static bool? FileFormatStandard(string filepath)
            {
                bool? valid = null;

                Process app = new Process();
                app.StartInfo.UseShellExecute = false;
                app.StartInfo.FileName = "javaw";
                string normal_dir = "C:\\Program Files\\ODF Validator\\odfvalidator-0.10.0-jar-with-dependencies.jar";

                // If app is run on Windows
                string? environ_dir = null;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    environ_dir = Environment.GetEnvironmentVariable("ODFValidator");
                }
                if (environ_dir != null)
                {
                    app.StartInfo.Arguments = $"-jar \"{environ_dir}\" \"{filepath}\"";
                }
                else
                {
                    app.StartInfo.Arguments = $"-jar \"{normal_dir}\" \"{filepath}\"";
                }

                app.Start();
                app.WaitForExit();
                int return_code = app.ExitCode;
                app.Close();

                if (return_code == 0)
                {
                    // File format is invalid. Spreadsheet has no cell values
                    valid = false;
                }
                if (return_code == 1)
                {
                    // File format validation could not be completed
                    valid = null;
                }
                if (return_code == 2)
                {
                    // File format is valid
                    valid = true;
                }
                return valid;
            }
        }
    }
}
