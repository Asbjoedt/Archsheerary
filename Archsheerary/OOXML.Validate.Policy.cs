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

                    bool valuesexist = check.ValuesExist(filepath);
                    List<Lists.FilePropertyInformation> filepropertyinformation = check.FilePropertyInformation(filepath);
                    bool conformance = check.Conformance(filepath);
                    List<Lists.DataConnections> connections = check.DataConnections(filepath);
                    List<Lists.ExternalCellReferences> extcellreferences = check.ExternalCellReferences(filepath);
                    List<Lists.RTDFunctions> rtdfunctions = check.RTDFunctions(filepath);
                    List<Lists.PrinterSettings> printersettings = check.PrinterSettings(filepath);
                    List<Lists.ExternalObjects> extobjects = check.ExternalObjects(filepath);
                    List<Lists.ActiveSheet> activesheet = check.ActiveSheet(filepath);
                    List<Lists.AbsolutePath> absolutepath = check.AbsolutePath(filepath);
                    List<Lists.EmbeddedObjects> embedobj = check.EmbeddedObjects(filepath);
                    List<Lists.Hyperlinks> hyperlinks = check.Hyperlinks(filepath);

                    // Add information to list and return it
                    results.Add(new Lists.OOXML.ValidatePolicy { ValuesExist = valuesexist, Conformance = conformance, DataConnections = connections, ExternalCellReferences = cellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, ActiveSheet = activesheet, AbsolutePath = absolutepath, EmbeddedObjects = embedobj, Hyperlinks = hyperlinks });
                    return results;
                }
            }
        }
    }
}
