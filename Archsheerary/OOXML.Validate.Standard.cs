using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.ComponentModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Archsheerary
{
    public partial class OOXML
    {
        public partial class Validate
        {
            public class Standard
            {
                /// <summary>
                /// Validate Office Open XML standard using Open XML SDK
                /// </summary>
                public List<DataTypes.OOXML.ValidateStandard> FileFormatStandard(string filepath)
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
                /// Validate OOXML file formats and ignore bug in Open XML SDK, which reports errors on Strict XLSX files
                /// </summary>
                public List<DataTypes.OOXML.ValidateStandard> FileFormatStandard_StrictHotfix(string filepath)
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
                            if (error_count <= 45)
                            {
                                foreach (var error in validation_errors)
                                {
                                    // Add validation results to list
                                    results.Add(new DataTypes.OOXML.ValidateStandard { IsValid = true, ErrorNumber = null, ErrorId = "", ErrorDescription = "", ErrorType = "", ErrorNode = "", ErrorPath = "", ErrorPart = "", ErrorRelatedNode = "", ErrorRelatedNodeInnerText = "" });
                                }
                            }
                            else
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
                        }
                        else
                        {
                            // Add validation results to list
                            results.Add(new DataTypes.OOXML.ValidateStandard { IsValid = true, ErrorNumber = null, ErrorId = "", ErrorDescription = "", ErrorType = "", ErrorNode = "", ErrorPath = "", ErrorPart = "", ErrorRelatedNode = "", ErrorRelatedNodeInnerText = "" });

                        }
                    }
                    return results;
                }
            }
        }
    }
}