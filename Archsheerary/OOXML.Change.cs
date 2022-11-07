﻿using System;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2013.ExcelAc;

namespace Archsheerary
{
    public partial class OOXML
    {
        public class Change
        {
            // Make first sheet active sheet
            public List<Lists.ActiveSheet> Activate_FirstSheet(string filepath)
            {
                List<Lists.ActiveSheet> results = new List<Lists.ActiveSheet>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                    WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                    if (workbookView.ActiveTab != null)
                    {
                        var activeSheetId = workbookView.ActiveTab.Value;

                        // Add to list
                        bool found = true;
                        results.Add(new Lists.ActiveSheet() { ActiveSheeet = activeSheetId, Found = found, Action = Lists.ActionChanged });

                        if (activeSheetId > 0)
                        {
                            // Set value in workbook.xml to first sheet
                            workbookView.ActiveTab.Value = 0;

                            // Iterate all worksheets to detect if sheetview.Tabselected exists and change it
                            IEnumerable<WorksheetPart> worksheets = spreadsheet.WorkbookPart.WorksheetParts;
                            foreach (WorksheetPart worksheet in worksheets)
                            {
                                SheetViews sheetviews = worksheet.Worksheet.SheetViews;
                                foreach (SheetView sheetview in sheetviews)
                                {
                                    sheetview.TabSelected = null;
                                }
                            }
                        }
                    }
                }
                return results;
            }
        }
    }
}
