using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public class Lists
    {
        public static string ActionChanged = "Changed";
        public static string ActionChecked = "Checked";
        public static string ActionRemoved = "Removed";

        public class DataConnections
        {
            public string? Id { get; set; }

            public string? Description { get; set; }

            public string? ConnectionFile { get; set; }

            public string? Credentials { get; set; }

            public string? DatabaseProperties { get; set; }

            public string Action { get; set; }
        }

        public class ExternalCellReferences
        {
            public string Sheet { get; set; }

            public string Cell { get; set; }

            public string Value { get; set; }

            public string Formula { get; set; }

            public string Target { get; set; }

            public string Action { get; set; }
        }

        public class RTDFunctions
        {
            public string Sheet { get; set; }

            public string Cell { get; set; }

            public string Value { get; set; }

            public string Formula { get; set; }

            public string? Action { get; set; }
        }

        public class ExternalObjects
        {
            public string Uri { get; set; }

            public string ContentType { get; set; }

            public string Target { get; set; }

            public string IsExternal { get; set; }

            public string Action { get; set; }
        }

        public class EmbeddedObjects
        {
            public string Uri { get; set; }

            public string ContentType { get; set; }

            public string Target { get; set; }

            public string IsExternal { get; set; }

            public string Action { get; set; }
        }

        public class PrinterSettings
        {
            public string Uri { get; set; }

            public string Action { get; set; }
        }

        public class AbsolutePath
        {
            public string Path { get; set; }

            public string Action { get; set; }
        }

        public class FilePropertyInformation
        {
            public string Author { get; set; }

            public string Title { get; set; }

            public string Keywords { get; set; }

            public string Category { get; set; }

            public string Subject { get; set; }

            public string Description { get; set; }

            public string LastModifiedBy { get; set; }

            public bool Found { get; set; } = false;

            public string Action { get; set; }
        }

        public class ActiveSheet
        {
            public uint ActiveSheeet { get; set; }

            public string? Action { get; set; }
        }

        public class Hyperlinks
        {
            public string Sheet { get; set; }

            public string Cell { get; set; }

            public string URL { get; set; }

            public string Action { get; set; }
        }

        public class ConformanceNamespaces
        {
            public string Prefix { get; set; }

            public string Transitional { get; set; }

            public string Strict { get; set; }
        }

        public class OOXML
        {
            public class ValidateStandard
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

            public class ValidatePolicy
            {
                public bool? ValuesExist { get; set; }

                public bool? Conformance { get; set; }

                public List<Lists.DataConnections> DataConnections { get; set; }

                public List<Lists.ExternalCellReferences> ExternalCellReferences { get; set; }

                public List<Lists.RTDFunctions> RTDFunctions { get; set; }

                public List<Lists.ExternalObjects> ExternalObjects { get; set; }

                public List<Lists.EmbeddedObjects> EmbeddedObjects { get; set; }

                public List<Lists.PrinterSettings> PrinterSettings { get; set; }

                public List<Lists.ActiveSheet> ActiveSheet { get; set; }

                public List<Lists.AbsolutePath> AbsolutePath { get; set; }

                public List<Lists.Hyperlinks> Hyperlinks { get; set; }

                public List<Lists.FilePropertyInformation> FilePropertyInformation { get; set; }
            }
        }

        public class FileFormatsIndex
        {
            public string Extension { get; set; }

            public string ExtensionUpper { get; set; }

            public string Description { get; set; }

            public string? Conformance { get; set; }

            public int? Count { get; set; }
        }

        public class OriginalFilesIndex
        {
            public string OriginalFilepath { get; set; }

            public string OriginalFilename { get; set; }

            public string OriginalExtension { get; set; }
        }

        public class FilesIndex
        {
            public string OriginalFilepath { get; set; }

            public string OriginalFilename { get; set; }

            public string OriginalExtension { get; set; }

            public string? NewFolderPath { get; set; }

            public string? CopyFilepath { get; set; }

            public string? CopyFilename { get; set; }

            public string? CopyExtension { get; set; }

            public string? ConversionFilepath { get; set; }

            public string? ConversionFilename { get; set; }

            public string? ConversionExtension { get; set; }

            public string? OOXMLConversionFilepath { get; set; }

            public string? OOXMLConversionFilename { get; set; }

            public string? OOXMLConversionExtension { get; set; }

            public string? ODSConversionFilepath { get; set; }

            public string? ODSConversionFilename { get; set; }

            public string? ODSConversionExtension { get; set; }

            public bool? ConversionSuccess { get; set; }
        }
    }
}
