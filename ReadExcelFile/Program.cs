// ====================================================================================================
//                                    Excel File Reader - Proof of Concept
// ====================================================================================================
// Programmed By:  Louiery R. Sincioco                                      Version:  .04 (Iteration 4)
// Programed Date:  October 23, 2019                                        Web Partners Group              
// ----------------------------------------------------------------------------------------------------
// Purpose:  Using .Net Core 3.0 and Open XML SDK (by Microsoft), read an Excel file which contains an
//           ImageLink column with a HYPERLINK formula that contains the URL to the product's image.
//           Every single product image is then downloaded to a specified location.  The program is
//           resilient enough that if it gets interrupted, it will resuming where it left off.
// ----------------------------------------------------------------------------------------------------
// Date           JIRA        Author     Description                                                   
// ----------------------------------------------------------------------------------------------------
// 10/23/2019     CDMC-30     Sin        Just read the Excel file cell-by-cell for general PoC.
// 10/23/2019     CDMC-31     Sin        Read only cells that contain the text "ImageLink"
// 10/24/2019     CDMC-4      Sin        Download the image that the "ImageLink" URL references.
// 10/24/2019     CDMC-33     Sin        Make program resilience and resume download if interrupted.
// ====================================================================================================

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace POC {

    static class ReadExcelFilePOC {

        // ------------------------------------------------------------------------------------------
        // Constants that may be stored in a configuration file or in a database
        // ------------------------------------------------------------------------------------------
        public const string TestExcelFile    = @".\Keystone Distributor Dometic Data with Images.xlsx";
        public const string DownloadLocation = @"C:\Temp\";
        public const string DownloadList     = DownloadLocation + "!ProductImageURLs.txt";
        public const string ExceptionList    = DownloadLocation + "!ProductImageURLs_FailedDownload.txt";
        public const string TempExtension    = ".wpg.tmp";

        // ------------------------------------------------------------------------------------------
        static void Main(string[] args) {

            Console.WriteLine("Sin's PoC for Reading an Excel Spreadsheet using .Net Core 3.0 and Open XML SDK.");
            Console.WriteLine("Copyright (c) 2019.  Web Partners Group.  All rights reserved.\n\n");

            WPG.Excel imageLinks = new WPG.Excel();

            // -----------------------------------------------------
            // Call our POC function - Iteration 1 - Read Cell Values
            // -----------------------------------------------------
            //WPG.POC.ReadExcelFileCellByCell(WPG.ReadExcelFilePOC.TestExcelFile);

            // -----------------------------------------------------
            // Call our POC function - Iteration 2 - Extract Image Links
            // -----------------------------------------------------
            List<string> URLs = new List<string>();

            // Extract Product Image URLs
            URLs = imageLinks.ExtractProductImageURLs(POC.ReadExcelFilePOC.TestExcelFile, "ImageLink");

            // Store the list into a file
            File.WriteAllLines(POC.ReadExcelFilePOC.DownloadList, URLs);

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

            // For Debugging - read in a list of URLs to download (the output of Iteration 2)
            //string[] arrImageList = System.IO.File.ReadAllLines(WPG.ReadExcelFilePOC.DownloadList);
            
            List<string> failedDownloadList = null;
            int SuccessfulDownloadCount = 0;

            SuccessfulDownloadCount = imageLinks.DownloadProductImages(URLs, out failedDownloadList);

            if (failedDownloadList != null && failedDownloadList.Count > 0) {

                // Save the list of files that failed to download
                System.IO.File.WriteAllLines(POC.ReadExcelFilePOC.ExceptionList, failedDownloadList);
            }

            Console.WriteLine("\nThere were {0} Product Image URLs and {1} were successfully downloaded.\n", URLs.Count, SuccessfulDownloadCount);
            Console.WriteLine("Press a key to continue...");
            Console.ReadKey();

        }
    }
}