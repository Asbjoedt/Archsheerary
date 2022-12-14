# Archive Spreadsheets Library

## What is Archsheerary
Archsheerary is a C# implementation of Open XML SDK and Excel Interop created as a library for purposes of digital archiving of spreadsheets. You can use the Archsheerary library to create your own applications for archiving of spreadsheets through workflows or single-use purposes.

* For more information, see repository **[CLISC](https://github.com/Asbjoedt/CLISC)**

## How to use
Install package in your project through [NuGet Gallery](https://www.nuget.org/packages/Archsheerary). You can then implement your own applications using Archsheerary through methods such as:

### Change
```
ExcelInterop.Change.ActivateFirstSheet()
ExcelInterop.Change.XLSXConformanceToTransitional()
ExcelInterop.Change.XLSXConformanceToStrict()
OOXML.Change.ActivateFirstSheet()
```
### Check
```
ExcelInterop.Check.ActiveSheet()
ExcelInterop.Check.DataConnections()
ExcelInterop.Check.ExternalCellReferences()
ExcelInterop.Check.FilePropertyInformation()
ExcelInterop.Check.RTDFunctions()
OOXML.Check.AbsolutePath()
OOXML.Check.ActiveSheet()
OOXML.Check.CellValues()
OOXML.Check.Conformance()
OOXML.Check.DataConnections()
OOXML.Check.EmbeddedObjects()
OOXML.Check.ExternalCellReferences()
OOXML.Check.ExternalObjects()
OOXML.Check.FilePropertyInformation()
OOXML.Check.Hyperlinks()
OOXML.Check.PrinterSettings()
OOXML.Check.RTDFunctions()

```
### Convert
```
ExcelInterop.ToAnySpreadsheetFileFormat()
ExcelInterop.ToODS()
ExcelInterop.ToXLSXTransitional()
ExcelInterop.ToXLSXStrict()
OOXML.Convert.ToXLSXTransitional()
OpenDocument.LibreOffice.ToAnySpreadsheetFileFormat()
OpenDocument.LibreOffice.ToODS()
OpenDocument.LibreOffice.ToXLSXTransitional()
```
### Remove
```
ExcelInterop.Remove.DataConnections()
ExcelInterop.Remove.ExternalCellReferences()
ExcelInterop.Remove.FilePropertyInformation()
ExcelInterop.Remove.RTDFunctions()
OOXML.Remove.AbsolutePath()
OOXML.Remove.DataConnections()
OOXML.Remove.ExternalCellReferences()
OOXML.Remove.ExternalObjects()
OOXML.Remove.EmbeddedObjects()
OOXML.Remove.Hyperlinks()
OOXML.Remove.PrinterSettings()
OOXML.Remove.RTDFunctions()
```
### Repair
```
OOXML.Repair()
```
### Validate
```
OOXML.Validate.Policy()
OOXML.Validate.Standard()
OpenDocument.Validate.Standard()
```
### Other
```
Other.Check.Extension()
Other.Checksum.MD5Hash()
Other.Compare.Spreadsheets()
Other.Count.Spreadsheets()
Other.Count.OOXMLConformance()
Other.Count.StrictConformance()
Other.Enumerate.Folder()
Other.Copy.Spreadsheet()
Other.FileFormats.FileFormatsIndex()
Other.FileFormats.ConformanceNamespacesIndex()
```

# Software & packages
The following software and packages are used under license.

* [Beyond Compare 4](https://www.scootersoftware.com/index.php), Copyright (c) 2022 Scooter Software, Inc.
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html), Apache License, [copyright info](https://github.com/tdf/odftoolkit/blob/master/NOTICE)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
