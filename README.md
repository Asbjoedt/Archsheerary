# Archive Spreadsheets Library (Archsheerary)

## What is Archsheerary
Archsheerary is a C# implementation of Open XML SDK and Excel Interop created as a library for purposes of digital archiving of spreadsheets. You can use the Archsheerary library to create your own applications for archiving of spreadsheets through workflows or single-use purposes.

## How to use Archsheerary
You can implement your own applications using Archsheerary through methods such as:

**Validate**
```
Validate.Standard()
Validate.Policy()
```
**Check**
```
Policy.Check.All()
Policy.Check.These()
Policy.Check.CellValues()
Policy.Check.Conformance()
Policy.Check.DataConnections()
Policy.Check.RTDFunctions()
Policy.Check.ExternalCellReferences()
Policy.Check.PrinterSettings()
Policy.Check.ExternalObjects()
Policy.Check.EmbeddedObjects()
Policy.Check.Hyperlinks()
Policy.Check.AbsolutePath()
```
**Change**
```
Policy.Change.All()
Policy.Change.These()
Policy.Change.Conformance()
Policy.Change.ActivateFirstSheet()
```
**Remove**
```
Policy.Remove.All()
Policy.Remove.These()
Policy.Remove.DataConnections()
Policy.Remove.RTDFunctions()
Policy.Remove.ExternalCellReferences()
Policy.Remove.ExternalObjects()
Policy.Remove.AbsolutePath()
Policy.Remove.EmbeddedObjects()
```

# Packages
The following packages are used under license.

* [Microsoft Excel Interop](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel), Copyright (c) Microsoft Corporation
* [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK), MIT License, Copyright (c) Microsoft Corporation
