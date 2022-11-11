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
                public static List<DataTypes.OOXML.ValidatePolicyAll> AllChecks(string filepath)
                {
                    List<DataTypes.OOXML.ValidatePolicyAll> results = new List<DataTypes.OOXML.ValidatePolicyAll>();

                    bool extension = Other.Check.Extension(filepath);
                    bool valuesexist = OOXML.Check.ValuesExist(filepath);
                    List<DataTypes.FilePropertyInformation> filepropertyinformation = OOXML.Check.FilePropertyInformation(filepath);
                    List<DataTypes.Conformance> conformance = OOXML.Check.Conformance(filepath);
                    List<DataTypes.DataConnections> connections = OOXML.Check.DataConnections(filepath);
                    List<DataTypes.ExternalCellReferences> extcellreferences = OOXML.Check.ExternalCellReferences(filepath);
                    List<DataTypes.RTDFunctions> rtdfunctions = OOXML.Check.RTDFunctions(filepath);
                    List<DataTypes.PrinterSettings> printersettings = OOXML.Check.PrinterSettings(filepath);
                    List<DataTypes.ExternalObjects> extobjects = OOXML.Check.ExternalObjects(filepath);
                    List<DataTypes.ActiveSheet> activesheet = OOXML.Check.ActiveSheet(filepath);
                    List<DataTypes.AbsolutePath> absolutepath = OOXML.Check.AbsolutePath(filepath);
                    List<DataTypes.EmbeddedObjects> embedobj = OOXML.Check.EmbeddedObjects(filepath);
                    List<DataTypes.Hyperlinks> hyperlinks = OOXML.Check.Hyperlinks(filepath);

                    // Add information to list and return it
                    results.Add(new DataTypes.OOXML.ValidatePolicyAll { Extension = extension, ValuesExist = valuesexist, FilePropertyInformation = filepropertyinformation, Conformance = conformance, DataConnections = connections, ExternalCellReferences = extcellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, ActiveSheet = activesheet, AbsolutePath = absolutepath, EmbeddedObjects = embedobj, Hyperlinks = hyperlinks });
                    return results;
                }

                /// <summary>
                /// Perform check of OPF specified preservation policy
                /// </summary>
                public static List<DataTypes.OOXML.ValidatePolicyOPF> OPFSpecification(string filepath)
                {
                    List<DataTypes.OOXML.ValidatePolicyOPF> results = new List<DataTypes.OOXML.ValidatePolicyOPF>();

                    bool extension = Other.Check.Extension(filepath);
                    bool valuesexist = OOXML.Check.ValuesExist(filepath);
                    List<DataTypes.Conformance> conformance = OOXML.Check.Conformance(filepath);
                    List<DataTypes.DataConnections> connections = OOXML.Check.DataConnections(filepath);
                    List<DataTypes.ExternalCellReferences> extcellreferences = OOXML.Check.ExternalCellReferences(filepath);
                    List<DataTypes.RTDFunctions> rtdfunctions = OOXML.Check.RTDFunctions(filepath);
                    List<DataTypes.PrinterSettings> printersettings = OOXML.Check.PrinterSettings(filepath);
                    List<DataTypes.ExternalObjects> extobjects = OOXML.Check.ExternalObjects(filepath);
                    List<DataTypes.AbsolutePath> absolutepath = OOXML.Check.AbsolutePath(filepath);
                    List<DataTypes.EmbeddedObjects> embedobj = OOXML.Check.EmbeddedObjects(filepath);

                    // Add information to list and return it
                    results.Add(new DataTypes.OOXML.ValidatePolicyOPF { Extension = extension, ValuesExist = valuesexist, Conformance = conformance, DataConnections = connections, ExternalCellReferences = extcellreferences, RTDFunctions = rtdfunctions, PrinterSettings = printersettings, ExternalObjects = extobjects, AbsolutePath = absolutepath, EmbeddedObjects = embedobj });
                    return results;
                }
            }
        }
    }
}
