using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    /// <summary>
    /// Collection of data types used by Archsheerary.
    /// </summary>
    public class DataTypes
    {
        public static string ActionChanged = "Changed";
        public static string ActionChecked = "Checked";
        public static string ActionRemoved = "Removed";

        public class Conformance
        {
            public string? OriginalConformance { get; set; }

            public string? NewConformance { get; set; }

            public string Action { get; set; }
        }
        public class DataConnections
        {
            public string? Id { get; set; }

            public string? Name { get; set; }

            public string? Description { get; set; }

            public string? Type { get; set; }

            public string? ConnectionFile { get; set; }

            public string? Credentials { get; set; }

            public string? DatabaseProperties { get; set; }

            public string? SourceFile { get; set; }

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

            public string Action { get; set; }
        }

        public class ExternalObjects
        {
            public string Target { get; set; }

            public string RelationshipType { get; set; }

            public bool IsExternal { get; set; }

            public string Container { get; set; }

            public string Action { get; set; }
        }

        public class EmbeddedObjects
        {
            public string Uri { get; set; }

            public string ContentType { get; set; }

            public string RelationshipType { get; set; }

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
            public uint? OriginalActiveSheet { get; set; }

            public uint? NewActiveSheet { get; set; }

            public string Action { get; set; }
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
                public bool? IsValid { get; set; }

                public int? ErrorNumber { get; set; }

                public string ErrorId { get; set; }

                public string ErrorDescription { get; set; }

                public string ErrorType { get; set; }

                public string ErrorNode { get; set; }

                public string ErrorPath { get; set; }

                public string ErrorPart { get; set; }

                public string? ErrorRelatedNode { get; set; }

                public string? ErrorRelatedNodeInnerText { get; set; }
            }

            public class ValidatePolicyAll
            {
                public bool Extension { get; set; }

                public bool? ValuesExist { get; set; }

                public List<DataTypes.Conformance> Conformance { get; set; }

                public List<DataTypes.DataConnections> DataConnections { get; set; }

                public List<DataTypes.ExternalCellReferences> ExternalCellReferences { get; set; }

                public List<DataTypes.RTDFunctions> RTDFunctions { get; set; }

                public List<DataTypes.ExternalObjects> ExternalObjects { get; set; }

                public List<DataTypes.EmbeddedObjects> EmbeddedObjects { get; set; }

                public List<DataTypes.PrinterSettings> PrinterSettings { get; set; }

                public List<DataTypes.ActiveSheet> ActiveSheet { get; set; }

                public List<DataTypes.AbsolutePath> AbsolutePath { get; set; }

                public List<DataTypes.Hyperlinks> Hyperlinks { get; set; }

                public List<DataTypes.FilePropertyInformation> FilePropertyInformation { get; set; }
            }

            public class ValidatePolicyOPF
            {
                public bool Extension { get; set; }

                public bool? ValuesExist { get; set; }

                public List<DataTypes.Conformance> Conformance { get; set; }

                public List<DataTypes.DataConnections> DataConnections { get; set; }

                public List<DataTypes.ExternalCellReferences> ExternalCellReferences { get; set; }

                public List<DataTypes.RTDFunctions> RTDFunctions { get; set; }

                public List<DataTypes.ExternalObjects> ExternalObjects { get; set; }

                public List<DataTypes.EmbeddedObjects> EmbeddedObjects { get; set; }

                public List<DataTypes.PrinterSettings> PrinterSettings { get; set; }

                public List<DataTypes.AbsolutePath> AbsolutePath { get; set; }
            }
        }

        public class FileFormatsIndex
        {
            public string Extension { get; set; }

            public string ExtensionUpper { get; set; }

            public string Description { get; set; }

            public string? Conformance { get; set; }

            public int? ExcelFileFormat { get; set; }

            public int? Count { get; set; }
        }

        public class OriginalFilesIndex
        {
            public string OriginalFilepath { get; set; }

            public string OriginalFilename { get; set; }

            public string OriginalExtension { get; set; }

            public string OriginalExtensionLower { get; set; }
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
