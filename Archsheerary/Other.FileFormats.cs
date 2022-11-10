using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Archsheerary
{
    public partial class Other
    {
        public class FileFormats
        {
            /// <summary>
            /// Create list of spreadsheet file formats
            /// </summary>
            public List<DataTypes.FileFormatsIndex> FileFormatsIndex()
            {
                List<DataTypes.FileFormatsIndex> list = new List<DataTypes.FileFormatsIndex>
                {
                    // GSHEET
                    new DataTypes.FileFormatsIndex {Extension = ".gsheet", ExtensionUpper = ".GSHEET", Description = "Google Sheets hyperlink"},
                    // FODS
                    new DataTypes.FileFormatsIndex {Extension = ".fods", ExtensionUpper = ".FODS", Description = "OpenDocument Flat XML Spreadsheet"},
                    // NUMBERS
                    new DataTypes.FileFormatsIndex {Extension = ".numbers", ExtensionUpper = ".NUMBERS", Description = "Apple Numbers Spreadsheet"},
                    // ODS
                    new DataTypes.FileFormatsIndex {Extension = ".ods", ExtensionUpper = ".ODS", Description = "OpenDocument Spreadsheet"},
                    // OTS
                    new DataTypes.FileFormatsIndex {Extension = ".ots", ExtensionUpper = ".OTS", Description = "OpenDocument Spreadsheet Template"},
                    // XLA
                    new DataTypes.FileFormatsIndex {Extension = ".xla", ExtensionUpper = ".XLA", Description = "Legacy Microsoft Excel Spreadsheet Add-In"},
                    // XLAM
                    new DataTypes.FileFormatsIndex {Extension = ".xlam", ExtensionUpper = ".XLAM", Description = "Office Open XML Macro-Enabled Add-In"},
                    // XLS
                    new DataTypes.FileFormatsIndex {Extension = ".xls", ExtensionUpper = ".XLS", Description = "Legacy Microsoft Excel Spreadsheet"},
                    // XLSB
                    new DataTypes.FileFormatsIndex {Extension = ".xlsb", ExtensionUpper = ".XLSB", Description = "Office Open XML Binary Spreadsheet"},
                    // XLSM
                    new DataTypes.FileFormatsIndex {Extension = ".xlsm", ExtensionUpper = ".XLSM", Description = "Office Open XML Macro-Enabled Spreadsheet"},
                    // XLSX - Transitional and Strict conformance
                    new DataTypes.FileFormatsIndex {Extension = ".xlsx", ExtensionUpper = ".XLSX", Description = "Office Open XML Spreadsheet (transitional and strict conformance)"},
                    // XLSX - Transitional conformance
                    new DataTypes.FileFormatsIndex {Extension = ".xlsx", ExtensionUpper = ".XLSX", Description = "Office Open XML Spreadsheet (transitional conformance)", Conformance = "transitional"},
                    // XLSX - Strict conformance
                    new DataTypes.FileFormatsIndex {Extension = ".xlsx", ExtensionUpper = ".XLSX", Description = "Office Open XML Spreadsheet (strict conformance)", Conformance = "strict"},
                    // XLT
                    new DataTypes.FileFormatsIndex {Extension = ".xlt", ExtensionUpper = ".XLT", Description = "Legacy Microsoft Excel Spreadsheet Template"},
                    // XLTM
                    new DataTypes.FileFormatsIndex {Extension = ".xltm", ExtensionUpper = ".XLTM", Description = "Office Open XML Macro-Enabled Spreadsheet Template"},
                    // XLTX
                    new DataTypes.FileFormatsIndex {Extension = ".XLTX", ExtensionUpper = ".XLTX", Description = "Office Open XML Spreadsheet Template"},
                };
                return list;
            }

            /// <summary>
            /// Create list of namespaces related to OOXML conformance
            /// </summary>
            public List<DataTypes.ConformanceNamespaces> ConformanceNamespacesIndex()
            {
                List<DataTypes.ConformanceNamespaces> list = new List<DataTypes.ConformanceNamespaces>();

                // xmlns
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "x", Transitional = "http://schemas.openxmlformats.org/spreadsheetml/2006/main", Strict = "http://purl.oclc.org/ooxml/spreadsheetml/main" });
                // docProps
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", Strict = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties" });
                // docProps/vt
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "vt", Transitional = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes", Strict = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" });
                // relationships/r
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "r", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships" });
                // relationship/styles
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles" });
                // relationship/theme
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme" });
                // relationship/worksheet
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" });
                // relationship/sharedStrings
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings" });
                // relationship/externalLink
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink" });
                // relationship/officeDocument
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" });
                // relationship/externallink/externalLinkPath
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLinkPath" });
                // relationship/hyperlink
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink" });
                // relationship/oleObject
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject" });
                // relationship/image
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/image" });
                // relationship/video
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/video" });
                // relationship/pivotCacheDefininition
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition" });
                // relationship/pivotCache Records
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheRecords" });
                // relationships/slicerCache
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.microsoft.com/office/2007/relationships/slicerCache", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/slicerCache" });
                // relationship/calcChain
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain" });
                // relationship/vmlDrawing - NO NAMESPACE FOR STRICT
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", Strict = "" });
                // relationship/drawing
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing" });
                // relationship/queryTable
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable" });
                // relationship/printerSettings
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings" });
                // relationship/comments
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments" });
                // relationship/vbaProject
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.microsoft.com/office/2006/relationships/vbaProject", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/vbaProject" });
                // relationship/xmlMaps
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps", Strict = "http://purl.oclc.org/ooxml/officeDocument/relationships/xmlMaps" });
                // drawingml/a
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "a", Transitional = "http://schemas.openxmlformats.org/drawingml/2006/main", Strict = "http://purl.oclc.org/ooxml/drawingml/main" });
                // drawingml/xdr
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "xdr", Transitional = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", Strict = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" });
                // drawingml/chart
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "c", Transitional = "http://schemas.openxmlformats.org/drawingml/2006/chart", Strict = "http://purl.oclc.org/ooxml/drawingml/chart" });
                // customXml/ds
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "ds", Transitional = "http://schemas.openxmlformats.org/officeDocument/2006/customXml", Strict = "" });
                // urn for Strict - NO NAMESPACE FOR TRANSITIONAL
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "v", Transitional = "", Strict = "urn:schemas-microsoft-com:vml" });
                // docProps/core.xml - NO NAMESPACE FOR TRANSITIONAL
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "dc", Transitional = "", Strict = "http://purl.org/dc/elements/1.1/" });
                // docProps/core.xml - NO NAMESPACE FOR TRANSITIONAL
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "dcterms", Transitional = "", Strict = "http://purl.org/dc/terms/" });
                // docProps/core.xml - NO NAMESPACE FOR TRANSITIONAL
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "dcmitype", Transitional = "", Strict = "http://purl.org/dc/dcmitype/" });
                // 
                list.Add(new DataTypes.ConformanceNamespaces() { Prefix = "", Transitional = "", Strict = "" });

                return list;
            }
        }
    }
}
