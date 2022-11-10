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
                /// <summary>
                /// Perform all available policy checks
                /// </summary>
                public List<Lists.OOXML.ValidatePolicy> AllChecks(string filepath)
                {
                    List<Lists.OOXML.ValidatePolicy> results = new List<Lists.OOXML.ValidatePolicy>();
                    OOXML.Check check = new OOXML.Check();

                    bool valuesexist = check.ValuesExist(filepath);
                    List<Lists.FilePropertyInformation> filepropertyinformation = check.FilePropertyInformation(filepath);
                    List<Lists.Conformance> conformance = check.Conformance(filepath);
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
                    results.Add(new Lists.OOXML.ValidatePolicy { ValuesExist = valuesexist, FilePropertyInformation = filepropertyinformation, Conformance = conformance, DataConnections = connections, ExternalCellReferences = extcellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, ActiveSheet = activesheet, AbsolutePath = absolutepath, EmbeddedObjects = embedobj, Hyperlinks = hyperlinks });
                    return results;
                }

                /// <summary>
                /// Perform check of OPF specified preservation policy
                /// </summary>
                public List<Lists.OOXML.ValidatePolicy> OPFSpecification(string filepath)
                {
                    List<Lists.OOXML.ValidatePolicy> results = new List<Lists.OOXML.ValidatePolicy>();
                    OOXML.Check check = new OOXML.Check();

                    bool valuesexist = check.ValuesExist(filepath);
                    bool conformance = check.Conformance(filepath);
                    List<Lists.DataConnections> connections = check.DataConnections(filepath);
                    List<Lists.ExternalCellReferences> extcellreferences = check.ExternalCellReferences(filepath);
                    List<Lists.RTDFunctions> rtdfunctions = check.RTDFunctions(filepath);
                    List<Lists.PrinterSettings> printersettings = check.PrinterSettings(filepath);
                    List<Lists.ExternalObjects> extobjects = check.ExternalObjects(filepath);
                    List<Lists.AbsolutePath> absolutepath = check.AbsolutePath(filepath);
                    List<Lists.EmbeddedObjects> embedobj = check.EmbeddedObjects(filepath);

                    // Add information to list and return it
                    results.Add(new Lists.OOXML.ValidatePolicy { ValuesExist = valuesexist, Conformance = conformance, DataConnections = connections, ExternalCellReferences = extcellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, AbsolutePath = absolutepath, EmbeddedObjects = embedobj });
                    return results;
                }
            }
        }
    }
}
