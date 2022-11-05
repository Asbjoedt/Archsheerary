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
                public List<ValidatePolicy> All(string filepath)
                {
                    List<ValidatePolicy> results = new List<ValidatePolicy>();
                    OOXML dp = new OOXML();

                    bool data = ValuesExist(filepath);
                    bool metadata = Metadata(filepath);
                    bool conformance = Conformance(filepath);
                    int connections = DataConnections(filepath);
                    int cellreferences = ExternalCellReferences(filepath);
                    int rtdfunctions = RTDFunctions(filepath);
                    int printersettings = PrinterSettings(filepath);
                    int extobjects = ExternalObjects(filepath);
                    bool activesheet = ActiveSheet(filepath);
                    bool absolutepath = AbsolutePath(filepath);
                    int embedobj = EmbeddedObjects(filepath);
                    int hyperlinks = Hyperlinks(filepath);

                    // Add information to list and return it
                    List<Check> Check = new List<Check>();
                    Check.Add(new Check { _ValuesExist = data, _Conformance = conformance, _DataConnections = connections, _ExternalCellReferences = cellreferences, _RTDFunctions = rtdfunctions, _PrinterSettings = printersettings, _ExternalObjects = extobjects, _ActiveSheet = activesheet, _AbsolutePath = absolutepath, _EmbeddedObjects = embedobj, _Hyperlinks = hyperlinks });
                    return Check;
                }
            }
        }
    }
}
