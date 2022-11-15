using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
        /// <summary>
        /// Collection of methods for validating Office Open XML spreadsheets.
        /// </summary>
        public partial class Validate
        {
            /// <summary>
            /// Validate Office Open XML standard using Open XML SDK. Returns list of errors.
            /// </summary>
            public static List<DataTypes.OOXML.ValidateStandard> FileFormatStandard(string filepath)
            {
                List<DataTypes.OOXML.ValidateStandard> results = new List<DataTypes.OOXML.ValidateStandard>();

                using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    var validator = new OpenXmlValidator();
                    var validation_errors = validator.Validate(spreadsheet).ToList();
                    int error_count = validation_errors.Count;
                    int error_number = 0;

                    if (validation_errors.Any()) // If errors
                    {
                        foreach (var error in validation_errors)
                        {
                            error_number++;

                            string er_rel_1 = "";
                            string er_rel_2 = "";
                            if (error.RelatedNode != null)
                            {
                                er_rel_1 = error.RelatedNode.ToString();
                                er_rel_2 = error.RelatedNode.InnerText;
                            }
                            // Add validation results to list
                            results.Add(new DataTypes.OOXML.ValidateStandard { IsValid = false, ErrorNumber = error_number, ErrorId = error.Id, ErrorDescription = error.Description, ErrorType = error.ErrorType.ToString(), ErrorNode = error.Node.ToString(), ErrorPath = error.Path.XPath.ToString(), ErrorPart = error.Part.Uri.ToString(), ErrorRelatedNode = er_rel_1, ErrorRelatedNodeInnerText = er_rel_2 });
                        }
                    }
                    else
                    {
                        // Add validation results to list
                        results.Add(new DataTypes.OOXML.ValidateStandard { IsValid = true, ErrorNumber = null, ErrorId = "", ErrorDescription = "", ErrorType = "", ErrorNode = "", ErrorPath = "", ErrorPart = "", ErrorRelatedNode = "", ErrorRelatedNodeInnerText = "" });
                    }
                }
                return results;
            }

            /// <summary>
            /// Validate OOXML file formats and ignore bug in Open XML SDK, which reports errors on Strict XLSX files. Returns list of errors.
            /// </summary>
            public static List<DataTypes.OOXML.ValidateStandard> FileFormatStandard_StrictHotfix(string filepath)
            {
                List<DataTypes.OOXML.ValidateStandard> results = new List<DataTypes.OOXML.ValidateStandard>();

                using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
                {
                    var validator = new OpenXmlValidator();
                    var validation_errors = validator.Validate(spreadsheet).ToList();
                    int error_count = validation_errors.Count;
                    int error_number = 0;

                    if (validation_errors.Any()) // If errors
                    {
                        foreach (var error in validation_errors)
                        {
                            // Open XML SDK has bugs, that is incorrectly reported as errors for Strict conformant spreadsheets. The switch suppresses these
                            switch (error.Id)
                            {
                                case "Sch_UndeclaredAttribute":
                                case "Sch_AttributeValueDataTypeDetailed":
                                    // Do nothing
                                    break;
                                default:
                                    error_number++;

                                    string er_rel_1 = "";
                                    string er_rel_2 = "";
                                    if (error.RelatedNode != null)
                                    {
                                        er_rel_1 = error.RelatedNode.ToString();
                                        er_rel_2 = error.RelatedNode.InnerText;
                                    }
                                    // Add validation results to list
                                    results.Add(new DataTypes.OOXML.ValidateStandard { IsValid = false, ErrorNumber = error_number, ErrorId = error.Id, ErrorDescription = error.Description, ErrorType = error.ErrorType.ToString(), ErrorNode = error.Node.ToString(), ErrorPath = error.Path.XPath.ToString(), ErrorPart = error.Part.Uri.ToString(), ErrorRelatedNode = er_rel_1, ErrorRelatedNodeInnerText = er_rel_2 });
                                    break;
                            }
                        }
                    }
                    else
                    {
                        // Add validation results to list
                        results.Add(new DataTypes.OOXML.ValidateStandard { IsValid = true, ErrorNumber = null, ErrorId = "", ErrorDescription = "", ErrorType = "", ErrorNode = "", ErrorPath = "", ErrorPart = "", ErrorRelatedNode = "", ErrorRelatedNodeInnerText = "" });

                    }
                }
                return results;
            }

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
