# Archive Spreadsheets Library

## What is Archsheerary
Archsheerary is a C# implementation of Open XML SDK and Excel Interop created as a library for purposes of digital archiving of spreadsheets. You can use the Archsheerary library to create your own applications for archiving of spreadsheets through workflows or single-use purposes.

* For more information, see repository **[CLISC](https://github.com/Asbjoedt/CLISC)**

## How to use
Install package in your project through [NuGet Gallery](https://www.nuget.org/packages/Archsheerary). You can then implement your own applications using Archsheerary through any of the below methods. Typical arguments are ```input_filepath```, ```output_filepath```, ```ouput_extension```, ```output_folder```, ```recurse``` and ```normalize```.

### Compare
```
BeyondCompare.Compare.Spreadsheets()
```
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
ExcelInterop.Convert.ToAnySpreadsheetFileFormat()
ExcelInterop.Convert.ToODS()
ExcelInterop.Convert.ToXLSXTransitional()
ExcelInterop.Convert.ToXLSXStrict()
LibreOffice.Convert.ToAnySpreadsheetFileFormat()
LibreOffice.Convert.ToODS()
LibreOffice.Convert.ToXLSXTransitional()
OOXML.Convert.EmbeddedImagesToTiff()
OOXML.Convert.ToXLSXTransitional()
```
### Extract
```
OOXML.Extract.EmbeddedObjects()
OOXML.Extract.ExternalObjects()
OOXML.Extract.FilePropertyInformation()
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
Other.Calculate.MD5Hash()
Other.Change.FileAttributesProtection()
Other.Check.ExtensionOOXMLAndOpenDocument()
Other.Check.ExtensionOOXML()
Other.Check.ExtensionOpenDocument()
Other.Check.FileAttributesProtection()
Other.Check.PasswordProtection()
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
* [Magick.Net](https://github.com/dlemstra/Magick.NET), Apache-2.0 license, copyright (c) Dirk Lemstra
* [Microsoft Excel](https://www.microsoft.com/en-us/microsoft-365/excel), Copyright (c) Microsoft Corporation
* [ODF Validator 0.10.0](https://odftoolkit.org/conformance/ODFValidator.html), Apache License, [copyright info](https://github.com/tdf/odftoolkit/blob/master/NOTICE)
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
