using System.IO;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Archsheerary
{
    public partial class ExcelInterop
    {
        public class Convert
        {
            // Convert any spreadsheet file format to another spreadsheet file format using Excel Interop
            public bool Perform(string input_filepath, string output_filepath, int output_fileformat)
            {
                bool convert_success = false;

                // Open Excel
                Excel.Application app = new Excel.Application(); // Create Excel object instance
                app.DisplayAlerts = false; // Don't display any Excel prompts
                Excel.Workbook wb = app.Workbooks.Open(input_filepath, ReadOnly: false, Password: "'", WriteResPassword: "'", IgnoreReadOnlyRecommended: true, Notify: false); // Create workbook instance

                // Save workbook and close Excel
                wb.SaveAs(output_filepath, output_fileformat);
                wb.Close();
                app.Quit();

                // If CLISC is run on Windows release Excel from task manager
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Marshal.ReleaseComObject(wb); // Delete workbook task in task manager
                    Marshal.ReleaseComObject(app); // Delete Excel task in task manager
                }

                // Repair spreadsheet
                Repair rep = new Repair();
                rep.Perform(output_filepath);

                // Return success
                convert_success = true;
                return convert_success;
            }
        }
    }
}
