using System;
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
    /// <summary>
    /// Collection of methods for Office Open XML spreadsheets.
    /// </summary>
    public partial class OOXML
    {
        /// <summary>
        /// Collection of methods for changing content in Office Open XML spreadsheets.
        /// </summary>
        public class Change
        {
            /// <summary>
            /// Make first sheet active sheet. Returns list of changed sheets.
            /// </summary>
            public static List<DataTypes.ActiveSheet> ActivateFirstSheet(string filepath)
            {
                List<DataTypes.ActiveSheet> results = new List<DataTypes.ActiveSheet>();

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, true))
                {
                    BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
                    WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
                    if (workbookView.ActiveTab != null)
                    {
                        if (workbookView.ActiveTab.Value > 0)
                        {
                            // Add to list
                            results.Add(new DataTypes.ActiveSheet() { OriginalActiveSheet = workbookView.ActiveTab.Value, NewActiveSheet = 0, Action = DataTypes.ActionChanged });

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
