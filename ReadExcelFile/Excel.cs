// ====================================================================================================
//                                          Extract Image URLs
// ====================================================================================================
// Programmed By:  Louiery R. Sincioco                                     Version:  1.0
// Programed Date:  October 24, 2019                                       Company:  Web Partners Group              
// ----------------------------------------------------------------------------------------------------
// Purpose:  Using .Net Core 3.0 and Open XML SDK (by Microsoft), read an Excel file which contains an
//           ImageLink column with a HYPERLINK formula that contains the URL to the product's image.
// ----------------------------------------------------------------------------------------------------
// Date           JIRA        Author     Description                                                   
// ----------------------------------------------------------------------------------------------------
// 10/23/2019     CDMC-31     Sin        Read only cells that contain the text "ImageLink"
// 10/24/2019     CDMC-4      Sin        Download the image that the "ImageLink" URL references.
// 10/24/2019     CDMC-33     Sin        Make program resilience and resume download if interrupted.
// ----------------------------------------------------------------------------------------------------
// Note:  This code is designed to read an Excel Spreadsheet with an "ImageLink" cell values
//        that contains a HYPERLINK formula as follows:
//
//        =HYPERLINK("http://Vehiclepartimages.com/pmdt/DMT/images/96010.jpg","ImageLink")
//
// ====================================================================================================

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WPG {

    public class Excel {

        // ------------------------------------------------------------------------------------------
        /// <summary>
        /// Returns the values of a specific cell value (for example 'ImageLink') in an Excel file.
        /// </summary>
        /// <param name="filename">The full path to where the Excel file is located</param>
        /// <param name="cellValueToLookFor">The cell value to look for (like 'ImageLink')</param>
        public List<string> ExtractProductImageURLs (string filename, string cellValueToLookFor) {

            List<string> result = new List<string>();

            // Check to ensure the file exists
            bool fileExists = File.Exists(filename);

            if (fileExists == true) {

                // Open the Excel file using Open XML SDK (a Microsoft library)
                using SpreadsheetDocument spreadhSheet = SpreadsheetDocument.Open(filename, false);

                // Reference the Workbooks
                WorkbookPart workbookPart = spreadhSheet.WorkbookPart;

                // Reference the Sheets collection
                Sheets sheetCollection = workbookPart.Workbook.GetFirstChild<Sheets>();

                // Loop through the Sheets collection
                foreach (Sheet sheet in sheetCollection.OfType<Sheet>()) {

                    // Reference a worksheet via its ID
                    Worksheet workSheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;

                    // Reference the data in the Sheet
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();

                    // Count the number of Rows
                    int rowCount = sheetData.ChildElements.Count();

                    // Iterate through the Rows
                    for (int row = 1; row < rowCount; row++) {

                        List<string> rowList = new List<string>();

                        // Count the number of Columns
                        int columnCount = sheetData.ElementAt(row).ChildElements.Count();

                        // Iterate through the Columns
                        for (int column = 0; column < columnCount; column++) {

                            string currentCellValue = string.Empty;

                            // Reference the current Cell
                            Cell currentCell = (Cell)sheetData.ElementAt(row).ChildElements.ElementAt(column);

                            // Are we in a cell with the Cell Value that we are looking for?
                            if (currentCell.CellValue.InnerText == cellValueToLookFor) {

                                // If so, extract the Cell Formula
                                string cellFormula = currentCell.CellFormula.InnerText;

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
            }

            return result;
        }

        // ------------------------------------------------------------------------------------------
        /// <summary>
        /// Given a list of URLs for product images, download them locally.  This operation can be
        /// stopped and resumed at any time in cases when we have Internet connection problems.
        /// </summary>
        /// <param name="URLs">A list of URLs</param>
        /// <param name="exceptionList">An output list of URL that failed to download</param>
        /// <param name="downloadLocation">The folder that where the images will be downloaded</param>
        /// <param name="temporaryExtension">The temporary extension added to the file being downloaded</param>
        /// <returns>Returns the number of images successfully downloaded</returns>
        public int DownloadProductImages (List<string> URLs, out List<string> exceptionList, string destinationFolder = @"C:\Temp\", string temporaryExtension = ".downloading") {

            // Initialize variables
            int downloadSuccessCount = 0;
            string downloadDestination = destinationFolder;
            exceptionList = new List<string>();

            // Add a trailing slash "\" if needed
            downloadDestination = destinationFolder.TrimEnd('\\') + @"\";

            // Initialize .Net's "internal" web browser / client
            using System.Net.WebClient wc = new System.Net.WebClient();

            // Iterate through our image list collection
            foreach (string URL in URLs) {

                string imageFileName = string.Empty;
                string imageFileNameTemp = string.Empty;

                // Step 1 - Extract just the file name portion of the Image Link (URL)
                Uri uri = new Uri(URL);
                imageFileName = Path.GetFileName(uri.LocalPath);

                // Step 2 - Download the Image
                if (String.IsNullOrEmpty(imageFileName) == false) {

                    Console.WriteLine("Downloading {0}", URL);

                    // Assign our real and temporary file names
                    imageFileName = destinationFolder + imageFileName;
                    imageFileNameTemp = imageFileName + temporaryExtension;

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
                        exceptionList.Add(URL);
                    }

                }
            }

            return downloadSuccessCount;
        }
    }
}
