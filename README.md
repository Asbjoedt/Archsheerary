# Archive Spreadsheets Library (Archsheerary)

## What is Archsheerary
Archsheerary is a C# implementation of Open XML SDK and Excel Interop created as a library for purposes of digital archiving of spreadsheets. You can use the Archsheerary library to create your own applications for archiving of spreadsheets through workflows or single-use purposes.

## How to use
You can implement your own applications using Archsheerary through methods such as:

**Change**
```
ExcelInterop.Change.Conformance()
OOXML.Change.ActivateFirstSheet()
```
**Check**
```
ExcelInterop.Check.Conformance()
OOXML.Check.CellValues()
OOXML.Check.DataConnections()
OOXML.Check.RTDFunctions()
OOXML.Check.ExternalCellReferences()
OOXML.Check.PrinterSettings()
OOXML.Check.ExternalObjects()
OOXML.Check.EmbeddedObjects()
OOXML.Check.Hyperlinks()
OOXML.Check.AbsolutePath()
```
**Convert**
```
ExcelInterop.Perform()
OOXML.Convert.ToXLSX()
```
**Remove**
```
OOXML.Remove.DataConnections()
OOXML.Remove.RTDFunctions()
OOXML.Remove.ExternalCellReferences()
OOXML.Remove.ExternalObjects()
OOXML.Remove.AbsolutePath()
OOXML.Remove.EmbeddedObjects()
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
Checksum.MD5()
Count.Spreadsheets()
```

# Packages
The following packages are used under license.

* [Microsoft Excel Interop](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel), Copyright (c) Microsoft Corporation
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
