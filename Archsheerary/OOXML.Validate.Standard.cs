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
                public List<Lists.OOXML.ValidateStandard> FileFormatStandard(string filepath)
                {
                    List<Lists.OOXML.ValidateStandard> results = new List<Lists.OOXML.ValidateStandard>();

                    using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
                    {
                        var validator = new OpenXmlValidator();
                        var validation_errors = validator.Validate(spreadsheet).ToList();
                        int error_count = validation_errors.Count;
                        int error_number = 0;

                        if (validation_errors.Any()) // If errors, return results
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
                                results.Add(new Lists.OOXML.ValidateStandard { Validity = "Invalid", Error_Number = error_number, Error_Id = error.Id, Error_Description = error.Description, Error_Type = error.ErrorType.ToString(), Error_Node = error.Node.ToString(), Error_Path = error.Path.XPath.ToString(), Error_Part = error.Part.Uri.ToString(), Error_RelatedNode = er_rel_1, Error_RelatedNode_InnerText = er_rel_2 });
                            }
                        }
                        else
                        {
                            // Add validation results to list
                            results.Add(new Lists.OOXML.ValidateStandard { Validity = "Valid", Error_Number = null, Error_Id = "", Error_Description = "", Error_Type = "", Error_Node = "", Error_Path = "", Error_Part = "", Error_RelatedNode = "", Error_RelatedNode_InnerText = "" });
                        }
                    }
                    return results;
                }

                // Validate Open Office XML file formats and ignoring bug in Open XML SDK, which reports errors on Strict .xlsx
                public List<Lists.OOXML.ValidateStandard> FileFormatStandard_StrictHotfix(string filepath)
                {
                    List<Lists.OOXML.ValidateStandard> results = new List<Lists.OOXML.ValidateStandard>();

                    using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
                    {
                        var validator = new OpenXmlValidator();
                        var validation_errors = validator.Validate(spreadsheet).ToList();
                        int error_count = validation_errors.Count;
                        int error_number = 0;

                        if (validation_errors.Any()) // If errors, return results
                        {
                            if (error_count >= 45)
                            {
                                foreach (var error in validation_errors)
                                {
                                    // Add validation results to list
                                    results.Add(new Lists.OOXML.ValidateStandard { Validity = "Valid", Error_Number = null, Error_Id = "", Error_Description = "", Error_Type = "", Error_Node = "", Error_Path = "", Error_Part = "", Error_RelatedNode = "", Error_RelatedNode_InnerText = "" });
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
                                            results.Add(new Lists.OOXML.ValidateStandard { Validity = "Invalid", Error_Number = error_number, Error_Id = error.Id, Error_Description = error.Description, Error_Type = error.ErrorType.ToString(), Error_Node = error.Node.ToString(), Error_Path = error.Path.XPath.ToString(), Error_Part = error.Part.Uri.ToString(), Error_RelatedNode = er_rel_1, Error_RelatedNode_InnerText = er_rel_2 });
                                            break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Add validation results to list
                            results.Add(new Lists.OOXML.ValidateStandard { Validity = "Valid", Error_Number = null, Error_Id = "", Error_Description = "", Error_Type = "", Error_Node = "", Error_Path = "", Error_Part = "", Error_RelatedNode = "", Error_RelatedNode_InnerText = "" });

                        }
                    }
                    return results;
                }
            }
        }
    }
}