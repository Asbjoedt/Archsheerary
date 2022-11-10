# Archive Spreadsheets Library

## What is Archsheerary
Archsheerary is a C# implementation of Open XML SDK and Excel Interop created as a library for purposes of digital archiving of spreadsheets. You can use the Archsheerary library to create your own applications for archiving of spreadsheets through workflows or single-use purposes.

## How to use
You can implement your own applications using Archsheerary through methods such as:

**Change**
```
ExcelInterop.Change.ActivateFirstSheet()
ExcelInterop.Change.XLSXConformanceToTransitional()
ExcelInterop.Change.XLSXConformanceToStrict()
OOXML.Change.ActivateFirstSheet()
```
**Check**
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
**Convert**
```
ExcelInterop.ToAnyFileFormat()
ExcelInterop.ToODS()
ExcelInterop.ToXLSXTransitional()
ExcelInterop.ToXLSXStrict()
OOXML.Convert.ToXLSXTransitioanl()
OpenDocument.LibreOffice.ToAnyFileFormat()
OpenDocument.LibreOffice.ToODS()
OpenDocument.LibreOffice.ToXLSXTransitional()
```
**Remove**
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
**Repair**
```
OOXML.Repair()
```
**Validate**
```
OOXML.Validate.Policy()
OOXML.Validate.Standard()
OpenDocument.Validate.Standard()
```
**Other**
```
Other.Checksum.MD5Hash()
Other.Compare.Spreadsheets()
Other.Count.Spreadsheets()
Other.Enumerate.Folder()
```

# Packages
The following packages are used under license.

* [Beyond Compare 4](https://www.scootersoftware.com/index.php), Copyright (c) 2022 Scooter Software, Inc.
* [LibreOffice](https://www.libreoffice.org/), Mozilla Public License v2.0
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html), Apache License, [copyright info](https://github.com/tdf/odftoolkit/blob/master/NOTICE)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
