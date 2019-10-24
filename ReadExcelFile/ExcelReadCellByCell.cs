// ====================================================================================================
//                                        Read Excel File Cell By Cell
// ====================================================================================================
// Programmed By:  Louiery R. Sincioco                                     Version:  1.0
// Programed Date:  October 24, 2019                                       Company:  Web Partners Group              
// ----------------------------------------------------------------------------------------------------
// Purpose:  Using .Net Core 3.0 and Open XML SDK (by Microsoft), read an Excel file cell value by
//           cell value.  This really was just a Proof of Concept (PoC).
// ----------------------------------------------------------------------------------------------------
// Date           JIRA        Author     Description                                                   
// ----------------------------------------------------------------------------------------------------
// 10/23/2019     CDMC-30     Sin        Just read the Excel file cell-by-cell for general PoC.
// ----------------------------------------------------------------------------------------------------
// Note:  The class is static as the methods are stateless.
// ====================================================================================================

using System;
using System.Linq;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace POC {

    static class Excel {

        // ------------------------------------------------------------------------------------------
        /// <summary>
        /// Outputs the content of all the cells in an Excel file on-screen.
        /// </summary>
        /// <param name="filename">The full path to where the Excel file is located</param>
        static public void ReadExcelFileCellByCell(string filename) {

            // Open the Excel file using Open XML SDK (a Microsoft library)
            using SpreadsheetDocument document = SpreadsheetDocument.Open(filename, false);

            // Reference the Workbook
            WorkbookPart workbookPart = document.WorkbookPart;

            // Reference the Sheets collection
            Sheets sheetCollection = workbookPart.Workbook.GetFirstChild<Sheets>();

            // Loop through the Sheets collection
            foreach (Sheet sheet in sheetCollection.OfType<Sheet>()) {

                // Reference a worksheet via its ID
                Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;

                // Reference the data in the Sheet
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                // Count the number of Rows
                int rowCount = sheetData.ChildElements.Count();

                // Iterate through the Rows
                for (int row = 0; row < rowCount; row++) {

                    List<string> rowList = new List<string>();

                    // Count the number of Columns
                    int columnCount = sheetData.ElementAt(row).ChildElements.Count();

                    // Iterate through the Columns
                    for (int column = 0; column < columnCount; column++) {

                        string currentCellValue = string.Empty;

                        // Reference the current Cell
                        Cell currentCell = (Cell)sheetData.ElementAt(row).ChildElements.ElementAt(column);

                        // Check if the Cell has data
                        if (currentCell.DataType != null) {

                            // If so, check if it is a Shared String (common string) like Column Headers
                            if (currentCell.DataType == CellValues.SharedString) {

                                int sharedStringID;

                                // ----------------------------------------------------------------------
                                // Internally, Excel creates a "normalized table" to store string values so they
                                // are not stored repeatedly within a Spreadsheet.  So you have to use the
                                // Shared String ID to get the text equivalent.
                                // ----------------------------------------------------------------------

                                // Let's see if we can parse a number out the Shared String ID
                                if (Int32.TryParse(currentCell.InnerText, out sharedStringID)) {

                                    // If we can, great, then let's turn that ID into its text equivalent
                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringID);

                                    // Check if we got a value
                                    if (item.Text != null) {

                                        // Are we on the first row?
                                        if (row == 0) {

                                            // If so, then we are probably just dealing with Column Headers
                                            Console.WriteLine(currentCell.CellReference + " (Text 1) = " + item.Text.Text);

                                        } else {

                                            // We are dealing with other Shared Strings that are not Column Headers
                                            Console.WriteLine(currentCell.CellReference + " (Text 2) = " + item.Text.Text);
                                        }
                                    }
                                }

                            } else {

                                // Check to see we are not dealing with Column Headers (just normal cell data)
                                if (row != 0) {

                                    // If so, then simply output the value (text) contained within the Cell
                                    Console.WriteLine(currentCell.CellReference + " (InnerText B) = " + currentCell.InnerText);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
