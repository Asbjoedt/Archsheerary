using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public class Lists
    {
        class DataConnections
        {
            string? _Description { get; set; }

            string? _Action { get; set; }
        }

        class ExternalCellReferences
        {
            string _Sheet { get; set; }

            string _Cell { get; set; }

            string _Value { get; set; }

            string _Formula { get; set; }

            string _Target { get; set; }

            string? _Action { get; set; }
        }

        class RTDFunctions
        {
            string _Sheet { get; set; }

            string _Cell { get; set; }

            string _Value { get; set; }

            string _Formula { get; set; }

            string? _Action { get; set; }
        }

        class ExternalObjects
        {
            string _Uri { get; set; }

            string _Target { get; set; }

            string _IsExternal { get; set; }

            bool? _Removed { get; set; }
        }

        class EmbeddedObjects
        {
            string? _Action { get; set; }
        }

        class PrinterSettings
        {
            string _Uri { get; set; }

            string? _Action { get; set; }
        }

        class AbsolutePath
        {
            bool _Null { get; set; }

            string _Path { get; set; }

            string? _Action { get; set; }
        }

        class FilePropertyInformation
        {
            string _Author { get; set; }

            string _Title { get; set; }

            string _Keyword { get; set; }

            string _LastModifiedBy { get; set; }

            string? _Action { get; set; }
        }

        class ActiveSheet
        {
            string _ActiveSheeet { get; set; }

            string? _Action { get; set; }
        }

        class Hyperlinks
        {
            string _Sheet { get; set; }

            string _Cell { get; set; }

            string _URL { get; set; }

            string? _Action { get; set; }
        }

        class ConformanceNamespaces
        {
            string _Prefix { get; set; }

            string _Transitional { get; set; }

            string _Strict { get; set; }
        }

        class OOXML
        {
            class ValidateStandard
            {
                public string Validity { get; set; }

                public int? Error_Number { get; set; }

                public string Error_Id { get; set; }

                public string Error_Description { get; set; }

                public string Error_Type { get; set; }

                public string Error_Node { get; set; }

                public string Error_Path { get; set; }

                public string Error_Part { get; set; }

                public string? Error_RelatedNode { get; set; }

                public string? Error_RelatedNode_InnerText { get; set; }
            }

            class ValidatePolicy
            {
                bool? _ValuesExist { get; set; }

                bool? _Conformance { get; set; }

                int? _DataConnections { get; set; }

                int? _ExternalCellReferences { get; set; }

                int? _RTDFunctions { get; set; }

                int? _ExternalObjects { get; set; }

                int? _EmbeddedObjects { get; set; }

                int? _PrinterSettings { get; set; }

                bool? _ActiveSheet { get; set; }

                bool? _FilePropertyInformation { get; set; }
            }
        }
    }
}
