// ====================================================================================================
//                                      Excel File Reader - Proof of Concept
// ====================================================================================================
// Programmed By:  Louiery R. Sincioco                                          Version:  .01 (Phase 1)
// Programed Date:  October 23, 2019                                                      
// ====================================================================================================
// Purpose:  Using .Net Core 3.0 and Open XML SDK (by Microsoft), see if we can read an Excel file
//           that Roberta has provided  which contains a ImageLink column with a URL to an image
//           we ultimately need to download and store in a network share.
// ====================================================================================================
// Phase 1:  Just read the Excel file cell-by-cell.
// Phase 2:  Read only the ImageLink column.
// Phase 3:  Download the image it references.

using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;

using Newtonsoft.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WPG {

    class ReadExcelFile {

        // ------------------------------------------------------------------------------------------
        static void Main(string[] args) {

            Console.WriteLine("Sin's PoC for Reading an Excel Spreadsheet using .Net Core 3.0 and Open XML SDK.");
            Console.WriteLine("Copyright (c) 2019.  Web Partners Group.  All rights reserved.\n\n");


            ReadExcelFileCellByCell(@"C:\WPG\Keystone Distributor Dometic Data with Images.xlsx");
            Console.ReadKey();
        }

        // ------------------------------------------------------------------------------------------
        static void ReadExcelFileCellByCell(string ExcelFile) {

            // Given an Excel file, open it and dump its content on screen one cell at a time.

            try {

                // Open the Excel file using Open XML SDK (a Microsoft library)
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ExcelFile, false)) {

                    // Reference the Workbooks
                    WorkbookPart workbookPart = doc.WorkbookPart;

                    // Reference the first Sheet in the Workbook
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();

                    // Loop through the Sheets collection
                    foreach (Sheet thesheet in thesheetcollection.OfType<Sheet>()) {

                        // Reference a worksheet via its ID
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        // Reference the data in the Sheet
                        SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();

                        // Count the number of Rows
                        int rowCount = thesheetdata.ChildElements.Count();

                        // Iterate through the Rows
                        for (int row = 0; row < rowCount; row++) {

                            List<string> rowList = new List<string>();

                            // Count the number of Columns
                            int columnCount = thesheetdata.ElementAt(row).ChildElements.Count();

                            // Iterate through the Columns
                            for (int column = 0; column < columnCount; column++) {

                                string currentcellvalue = string.Empty;

                                // Reference the current Cell
                                Cell thecurrentcell = (Cell)thesheetdata.ElementAt(row).ChildElements.ElementAt(column);

                                // Check if the Cell has data
                                if (thecurrentcell.DataType != null) {

                                    // If so, check if it is a Shared String (common string) like Column Headers
                                    if (thecurrentcell.DataType == CellValues.SharedString) {

                                        int sharedStringID;

                                        // Internally, Excel creates a "normalized table" to store string values so they
                                        // are not stored repeatedly within a Spreadsheet.

                                        // Let's see if we can parse a number out the Shared String ID
                                        if (Int32.TryParse(thecurrentcell.InnerText, out sharedStringID)) {

                                            // If we can, great, then let's turn that ID into its text equivalent
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringID);

                                            // Check if we got a value
                                            if (item.Text != null) {

                                                // Are we on the first row?
                                                if (row == 0) {
                                                    //dtTable.Columns.Add(item.Text.Text);

                                                    // If so, then we are probably just dealing with Column Headers
                                                    Console.WriteLine(thecurrentcell.CellReference + " (Text 1) = " + item.Text.Text);
                                                } else {
                                                    //rowList.Add(item.Text.Text);

                                                    // We are dealing with other Shared Strings that are not Column Headers
                                                    Console.WriteLine(thecurrentcell.CellReference + " (Text 2) = " + item.Text.Text);
                                                }
                                            }
                                        } else if (thecurrentcell.DataType == CellValues.String) {
                                            //Console.WriteLine(thecurrentcell.InnerText);
                                            Console.WriteLine(thecurrentcell.CellReference + " (InnerText A) = " + thecurrentcell.InnerText);
                                        }

                                    }
                                } else {

                                    // Check to see we are not dealing with Column Headers
                                    if (row != 0) {
                                        //rowList.Add(thecurrentcell.InnerText);

                                        // If so, then simply output the value (text) contained within the Cell
                                        Console.WriteLine(thecurrentcell.CellReference + " (InnerText B) = " + thecurrentcell.InnerText);
                                    }

                                }
                            }
                            // if (row != 0) {
                            //reserved for column values
                            //dtTable.Rows.Add(rowList.ToArray());
                            //}

                        }

                    }

                    //return JsonConvert.SerializeObject(dtTable);
                    //return string.Empty;
                }
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        // ------------------------------------------------------------------------------------------
        static void TempCode(string ExcelFile) {

            // Open the document for editing.
            //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(ExcelFile, false))
            //{

            //	WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            //	WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            //	OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            //	string text;
            //	while (reader.Read())
            //	{
            //		Console.WriteLine(reader.ElementType);


            //		if (reader.ElementType == typeof(CellValue))
            //		{
            //			text = reader.GetText();
            //			Console.Write(text + " ");
            //		}
            //	}

            //	return string.Empty;
            //}

            //using (ExcelPackage package = new ExcelPackage(ExcelFile))
            //{
            //	ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
            //	int rowCount = worksheet.Dimension.Rows;
            //	int ColCount = worksheet.Dimension.Columns;

            //	var rawText = string.Empty;
            //	for (int row = 1; row <= rowCount; row++)
            //	{
            //		for (int col = 1; col <= ColCount; col++)
            //		{
            //			// This is just for demo purposes
            //			rawText += worksheet.Cells[row, col].Value.ToString() + "\t";
            //		}
            //		rawText += "\r\n";
            //	}
            //	_logger.LogInformation(rawText);
            //}


        }
    }
}
