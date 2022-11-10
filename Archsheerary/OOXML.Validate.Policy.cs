using DocumentFormat.OpenXml.Wordprocessing;
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
                public List<DataTypes.OOXML.ValidatePolicyAll> AllChecks(string filepath)
                {
                    List<DataTypes.OOXML.ValidatePolicyAll> results = new List<DataTypes.OOXML.ValidatePolicyAll>();
                    OOXML.Check check = new OOXML.Check();
                    Other.Check check2 = new Other.Check();

                    bool extension = check2.Extension(filepath);
                    bool valuesexist = check.ValuesExist(filepath);
                    List<DataTypes.FilePropertyInformation> filepropertyinformation = check.FilePropertyInformation(filepath);
                    List<DataTypes.Conformance> conformance = check.Conformance(filepath);
                    List<DataTypes.DataConnections> connections = check.DataConnections(filepath);
                    List<DataTypes.ExternalCellReferences> extcellreferences = check.ExternalCellReferences(filepath);
                    List<DataTypes.RTDFunctions> rtdfunctions = check.RTDFunctions(filepath);
                    List<DataTypes.PrinterSettings> printersettings = check.PrinterSettings(filepath);
                    List<DataTypes.ExternalObjects> extobjects = check.ExternalObjects(filepath);
                    List<DataTypes.ActiveSheet> activesheet = check.ActiveSheet(filepath);
                    List<DataTypes.AbsolutePath> absolutepath = check.AbsolutePath(filepath);
                    List<DataTypes.EmbeddedObjects> embedobj = check.EmbeddedObjects(filepath);
                    List<DataTypes.Hyperlinks> hyperlinks = check.Hyperlinks(filepath);

                    // Add information to list and return it
                    results.Add(new DataTypes.OOXML.ValidatePolicyAll { Extension = extension, ValuesExist = valuesexist, FilePropertyInformation = filepropertyinformation, Conformance = conformance, DataConnections = connections, ExternalCellReferences = extcellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, ActiveSheet = activesheet, AbsolutePath = absolutepath, EmbeddedObjects = embedobj, Hyperlinks = hyperlinks });
                    return results;
                }

                /// <summary>
                /// Perform check of OPF specified preservation policy
                /// </summary>
                public List<DataTypes.OOXML.ValidatePolicyOPF> OPFSpecification(string filepath)
                {
                    List<DataTypes.OOXML.ValidatePolicyOPF> results = new List<DataTypes.OOXML.ValidatePolicyOPF>();
                    OOXML.Check check = new OOXML.Check();
                    Other.Check check2 = new Other.Check();

                    bool extension = check2.Extension(filepath);
                    bool valuesexist = check.ValuesExist(filepath);
                    List<DataTypes.Conformance> conformance = check.Conformance(filepath);
                    List<DataTypes.DataConnections> connections = check.DataConnections(filepath);
                    List<DataTypes.ExternalCellReferences> extcellreferences = check.ExternalCellReferences(filepath);
                    List<DataTypes.RTDFunctions> rtdfunctions = check.RTDFunctions(filepath);
                    List<DataTypes.PrinterSettings> printersettings = check.PrinterSettings(filepath);
                    List<DataTypes.ExternalObjects> extobjects = check.ExternalObjects(filepath);
                    List<DataTypes.AbsolutePath> absolutepath = check.AbsolutePath(filepath);
                    List<DataTypes.EmbeddedObjects> embedobj = check.EmbeddedObjects(filepath);

                    // Add information to list and return it
                    results.Add(new DataTypes.OOXML.ValidatePolicyOPF { Extension = extension, ValuesExist = valuesexist, Conformance = conformance, DataConnections = connections, ExternalCellReferences = extcellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, AbsolutePath = absolutepath, EmbeddedObjects = embedobj });
                    return results;
                }
            }
        }
    }
}
