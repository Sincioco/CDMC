// ====================================================================================================
//                                      Excel File Reader - Proof of Concept
// ====================================================================================================
// Programmed By:  Louiery R. Sincioco                                      Version:  .04 (Iteration 4)
// Programed Date:  October 24, 2019                                                      
// ----------------------------------------------------------------------------------------------------
// Purpose:  Using .Net Core 3.0 and Open XML SDK (by Microsoft), see if we can read an Excel file
//           that Roberta has provided  which contains a ImageLink column with a URL to an image
//           we ultimately need to download and store in a network share.
// ----------------------------------------------------------------------------------------------------
// Iteration 1:  Just read the Excel file cell-by-cell.             -COMPLETED
// Iteration 2:  Read only the ImageLink column.                    -COMPLETED
// Iteration 3:  Download the image it references.                  -COMPLETED
// Iteration 4:  Add error recovery so we can resume download.      -COMPLETED
// ====================================================================================================

using System;
using System.Collections.Generic;
using System.Linq;

using Newtonsoft.Json;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace WPG {

    static class ReadExcelFilePOC {

        public const string TestExcelFile = @"C:\WPG\Keystone Distributor Dometic Data with Images.xlsx";
        public const string DownloadLocation = @"C:\Temp\";
        public const string DownloadList = @"C:\Temp\!ProductImageURLs.txt";
        public const string ExceptionList = @"C:\Temp\!ProductImageURLs_FailedDownload.txt";
        public const string DownloadTempFileExtension = ".wpgdownload.tmp";

        // ------------------------------------------------------------------------------------------
        static void Main(string[] args) {

            Console.WriteLine("Sin's PoC for Reading an Excel Spreadsheet using .Net Core 3.0 and Open XML SDK.");
            Console.WriteLine("Copyright (c) 2019.  Web Partners Group.  All rights reserved.\n\n");

            // -----------------------------------------------------
            // Call our POC function - Iteration 1 - Read Cell Values
            // -----------------------------------------------------
            //ReadExcelFileCellByCell(WPG.ReadExcelFilePOC.TestExcelFile);

            // -----------------------------------------------------
            // Call our POC function - Iteration 2 - Extract Image Links
            // -----------------------------------------------------
            List<string> URLs = new List<string>();

            // Extract Product Image URLs
            URLs = ExtractProductImageURLs(WPG.ReadExcelFilePOC.TestExcelFile, "ImageLink");

            // Store the list into a file
            File.WriteAllLines(WPG.ReadExcelFilePOC.DownloadList, URLs);

            // Ourput each Product URL on-screen
            foreach (string URL in URLs) {
                Console.WriteLine(URL);
            }

            Console.WriteLine("\nYour Excel File contained {0} Product Image URLs\n", URLs.Count);

            Console.WriteLine("Press a key to continue...");
            Console.ReadKey();

            // -----------------------------------------------------
            // Call our POC function - Iteration 3 & 4 - Save Images - Allow for Resume/Recovery
            // -----------------------------------------------------
            string[] arrImageList = System.IO.File.ReadAllLines(WPG.ReadExcelFilePOC.DownloadList);
            
            List<string> FailedDownloadList = null;
            int SuccessfullDownloadCount = 0;

            SuccessfullDownloadCount = DownloadImages(URLs, out FailedDownloadList);

            if (FailedDownloadList != null && FailedDownloadList.Count > 0) {

                // Save the list of files that failed to download
                System.IO.File.WriteAllLines(WPG.ReadExcelFilePOC.ExceptionList, FailedDownloadList);
            }

            Console.WriteLine("\nThere were {0} Product Image URLs and {1} were successfully downloaded.\n", URLs.Count, SuccessfullDownloadCount);

            Console.WriteLine("Press a key to continue...");
            Console.ReadKey();

        }

        // ------------------------------------------------------------------------------------------
        /// <summary>
        /// Outputs the content of all the cells in an Excel file on-screen.
        /// </summary>
        /// <param name="filename">The full path to where the Excel file is located</param>
        static void ReadExcelFileCellByCell(string filename) {

            // Open the Excel file using Open XML SDK (a Microsoft library)
            using SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false);

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

                                // ----------------------------------------------------------------------
                                // Internally, Excel creates a "normalized table" to store string values so they
                                // are not stored repeatedly within a Spreadsheet.
                                // ----------------------------------------------------------------------

                                // Let's see if we can parse a number out the Shared String ID
                                if (Int32.TryParse(thecurrentcell.InnerText, out sharedStringID)) {

                                    // If we can, great, then let's turn that ID into its text equivalent
                                    SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringID);

                                    // Check if we got a value
                                    if (item.Text != null) {

                                        // Are we on the first row?
                                        if (row == 0) {

                                            // If so, then we are probably just dealing with Column Headers
                                            Console.WriteLine(thecurrentcell.CellReference + " (Text 1) = " + item.Text.Text);

                                        } else {

                                            // We are dealing with other Shared Strings that are not Column Headers
                                            Console.WriteLine(thecurrentcell.CellReference + " (Text 2) = " + item.Text.Text);
                                        }
                                    }
                                }

                            } else {

                                // Check to see we are not dealing with Column Headers (just normal cell data)
                                if (row != 0) {

                                    // If so, then simply output the value (text) contained within the Cell
                                    Console.WriteLine(thecurrentcell.CellReference + " (InnerText B) = " + thecurrentcell.InnerText);
                                }
                            }
                        }
                    }
                }
            }
        }

        // ------------------------------------------------------------------------------------------
        /// <summary>
        /// Returns the values of a specific column in an Excel file.
        /// </summary>
        /// <param name="filename">The full path to where the Excel file is located</param>
        /// <param name="CellValueToLookFor">The cell value to look for (like 'ImageLink')</param>
        static List<string> ExtractProductImageURLs(string filename, string CellValueToLookFor) {

            List<string> result = new List<string>();

            // Open the Excel file using Open XML SDK (a Microsoft library)
            using SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false);

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
                for (int row = 1; row < rowCount; row++) {

                    List<string> rowList = new List<string>();

                    // Count the number of Columns
                    int columnCount = thesheetdata.ElementAt(row).ChildElements.Count();

                    // Iterate through the Columns
                    for (int column = 0; column < columnCount; column++) {

                        string currentcellvalue = string.Empty;

                        // Reference the current Cell
                        Cell thecurrentcell = (Cell)thesheetdata.ElementAt(row).ChildElements.ElementAt(column);

                        // Are we in a cell with the Cell Value that we are looking for?
                        if (thecurrentcell.CellValue.InnerText == CellValueToLookFor) {

                            // If so, extract the Cell Formula
                            string cellFormula = thecurrentcell.CellFormula.InnerText;

                            // ----------------------------------------------------------------------
                            // The Cell Formula is in this format:
                            // =HYPERLINK("http://Vehiclepartimages.com/pmdt/DMT/images/96010.jpg","ImageLink")
                            // ----------------------------------------------------------------------

                            // Split it by the double quotes
                            string[] arrFormula = cellFormula.Split("\"");

                            if (arrFormula.Length > 0) {

                                // The URL itself will be in the second index
                                string URL = arrFormula[1];

                                // Store the URL in our result accumulator
                                result.Add(URL);
                            }
                        }
                    }
                }
            }

            return result;
        }

        // ------------------------------------------------------------------------------------------
        /// <summary>
        /// Given a list of URLs for product images, download them locally.  This operation can be
        /// stopped and resumed at any time in cases when we have Internet connection problems.
        /// </summary>
        /// <param name="ImageList">A list of URLs</param>
        /// <param name="FailedDownloadList">An output list of URL (images) we failed to download</param>
        /// <returns>Returns the number of images successfully downloaded</returns>
        static int DownloadImages(List<string> ImageList, out List<string> FailedDownloadList) {

            int downloadSuccessCount = 0;
            FailedDownloadList = new List<string>();

            // Initialize .Net's "internal" web browser / client
            using System.Net.WebClient wc = new System.Net.WebClient();

            // Iterate through our image list collection
            foreach (string URL in ImageList) {

                string imageFileName = string.Empty;
                string imageFileNameTemp = string.Empty;

                // Step 1 - Extract just the file name portion of the Image Link (URL)
                Uri uri = new Uri(URL);
                imageFileName = Path.GetFileName(uri.LocalPath);

                // Step 2 - Download the Image
                if (String.IsNullOrEmpty(imageFileName) == false) {
                    
                    Console.WriteLine("Downloading {0}", URL);

                    // Assign our real and temporary file names
                    imageFileName = WPG.ReadExcelFilePOC.DownloadLocation + imageFileName;
                    imageFileNameTemp = imageFileName + WPG.ReadExcelFilePOC.DownloadTempFileExtension;

                    try {

                        // Check if we have already downloaded the file previously
                        if (File.Exists(imageFileName) == false) {

                            // If not, download the image using the temporary file name
                            wc.DownloadFile(URL, imageFileNameTemp);

                            // Rename it to the real file name after download
                            System.IO.File.Move(imageFileNameTemp, imageFileName);
                        }

                        // Increment our download success count
                        downloadSuccessCount++;

                    } catch (Exception) {

                        Console.WriteLine("\tFailed to download {0}", URL);

                        // Remember Image URLs that failed to download
                        FailedDownloadList.Add(URL);
                    }
                    
                }
            }

            return downloadSuccessCount;

        }
    }
}