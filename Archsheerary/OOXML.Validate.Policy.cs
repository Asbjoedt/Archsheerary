using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public partial class OOXML
    {
        public partial class Validate
        {
            public class Policy
            {
                // Perform check of archival requirements
                public List<Lists.OOXML.ValidatePolicy> All(string filepath)
                {
                    List<Lists.OOXML.ValidatePolicy> results = new List<Lists.OOXML.ValidatePolicy>();
                    OOXML.Check check = new OOXML.Check();

                    bool data = check.ValuesExist(filepath);
                    bool metadata = check.FilePropertyInformation(filepath);
                    bool conformance = check.Conformance(filepath);
                    int connections = check.DataConnections(filepath);
                    int cellreferences = check.ExternalCellReferences(filepath);
                    int rtdfunctions = check.RTDFunctions(filepath);
                    int printersettings = check.PrinterSettings(filepath);
                    int extobjects = check.ExternalObjects(filepath);
                    bool activesheet = check.ActiveSheet(filepath);
                    bool absolutepath = check.AbsolutePath(filepath);
                    int embedobj = check.EmbeddedObjects(filepath);
                    int hyperlinks = check.Hyperlinks(filepath);

                    // Add information to list and return it
                    results.Add(new Lists.OOXML.ValidatePolicy { ValuesExist = data, Conformance = conformance, DataConnections = connections, ExternalCellReferences = cellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, ActiveSheet = activesheet, AbsolutePath = absolutepath, EmbeddedObjects = embedobj, Hyperlinks = hyperlinks });
                    return results;
                }
            }
        }
    }
}
