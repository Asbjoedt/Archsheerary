using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    /// <summary>
    /// Collection of other metods related to archiving spreadsheets.
    /// </summary>
    public partial class Other
    {
        /// <summary>
        /// Collection of methods for checking validity of spreadsheets.
        /// </summary>
        public class Check
        {
            /// <summary>
            /// Check for OOXML and OpenDocument file format extensions. Returns true boolean if valid extension.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="normalize">True if only .xlsx and .ods extensions are valid</param>
            /// <return>True bool if extension is valid</return>
            public static bool ExtensionOOXMLAndOpenDocument(string filepath, bool normalize)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (!normalize)
                {
                    if (extension == ".xlam" || extension == ".xlsx" || extension == ".xlsm" || extension == ".xltm" || extension == ".xlsx" || extension == ".fods" || extension == ".ods" || extension == ".ots")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                else if (normalize)
                {
                    if (extension == ".xlsx" || extension == ".ods")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                return validextension;
            }

            /// <summary>
            /// Check for OOXML extensions. Returns true boolean if valid extension.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="normalize">True if only .xlsx extension is valid</param>
            /// <return>True bool if extension is valid</return>
            public static bool ExtensionOOXML(string filepath, bool normalize)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (!normalize)
                {
                    if (extension == ".xlam" || extension == ".xlsx" || extension == ".xlsm" || extension == ".xltm" || extension == ".xlsx")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                else if (normalize)
                {
                    if (extension == ".xlsx")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                return validextension;
            }

            /// <summary>
            /// Check for OpenDocument Spreadsheets extensions. Returns true boolean if valid extension.
            /// </summary>
            /// <param name="filepath">Path to input file</param>
            /// <param name="normalize">True if only .ods extension is valid</param>
            /// <return>True bool if extension is valid</return>
            public static bool ExtensionOpenDocument(string filepath, bool normalize)
            {
                bool validextension = false;

                FileInfo file_info = new FileInfo(filepath);
                string extension = file_info.Extension.ToLower();
                if (!normalize)
                {
                    if (extension == ".fods" || extension == ".ods" || extension == ".ots")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                else if (normalize)
                {
                    if (extension == ".ods")
                    {
                        validextension = true;
                        return validextension;
                    }
                }
                return validextension;
            }
        }
    }
}
