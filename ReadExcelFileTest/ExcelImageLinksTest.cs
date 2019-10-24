// ====================================================================================================
//                         Unit Testing for Product Image Extraction and Download
// ====================================================================================================
// Programmed By:  Louiery R. Sincioco                                     Version:  1.0
// Programed Date:  October 24, 2019                                       Company:  Web Partners Group              
// ----------------------------------------------------------------------------------------------------
// Purpose:  Create unit testing to ensure that we can extract URLs from an Excel file and download
//           the images that those URLs references.
// ----------------------------------------------------------------------------------------------------
// Date           JIRA        Author     Description                                                   
// ----------------------------------------------------------------------------------------------------
// 10/23/2019     CDMC-31     Sin        Read only cells that contain the text "ImageLink"
// 10/24/2019     CDMC-4      Sin        Download the image that the "ImageLink" URL references.
// 10/24/2019     CDMC-33     Sin        Make program resilience and resume download if interrupted.
// ----------------------------------------------------------------------------------------------------
// Note:  This unit test file will run two tests:
//        1.) I will extract all the Product Image URLs from an Excel Spreadsheet.
//        2.) It will download all of them.
// ====================================================================================================

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace WPG {

    [TestClass]
    public class ExcelImageLinksTest {

        string TestExcelFile = @".\Keystone Distributor Dometic Data with Images.xlsx";
        string DownloadList = @"C:\Temp\!ProductImageURLs.txt";
        string ExceptionList = @"C:\Temp\!ProductImageURLs_FailedDownload.txt";

        [TestMethod]
        public void T1_ExtractProductImageURLs() {

            string currDirectory = Directory.GetCurrentDirectory();

            // ----------------------------------------------------------------------
            // Arange
            // ----------------------------------------------------------------------            
            Excel imageLinks = new Excel();

            // ----------------------------------------------------------------------
            // Act
            // ----------------------------------------------------------------------

            // Extract Product Image URLs from the Spreedsheet
            List<string> URLs = imageLinks.ExtractProductImageURLs(this.TestExcelFile, "ImageLink");

            // Save the list of extracted URLs
            File.WriteAllLines(this.DownloadList, URLs);

            // ----------------------------------------------------------------------
            // Assert
            // ----------------------------------------------------------------------
            Assert.IsTrue(URLs.Count == 247, "Test Passed: We extracted the correct number of URLs.");

        }

        [TestMethod]
        public void T2_DownloadProductImage() {

            // ----------------------------------------------------------------------
            // Arange
            // ----------------------------------------------------------------------            
            int SuccessfulDownloadCount = 0;
            List<string> failedDownloadList = new List<string>();

            Excel imageLinks = new Excel();
            List<string> URLs = imageLinks.ExtractProductImageURLs(this.TestExcelFile, "ImageLink");
            
            // ----------------------------------------------------------------------
            // Act
            // ----------------------------------------------------------------------
            SuccessfulDownloadCount = imageLinks.DownloadProductImages(URLs, out failedDownloadList);

            if (failedDownloadList != null && failedDownloadList.Count > 0) {

                // Save the list of files that failed to download
                System.IO.File.WriteAllLines(this.ExceptionList, failedDownloadList);
            }

            // ----------------------------------------------------------------------
            // Assert
            // ----------------------------------------------------------------------
            Assert.IsTrue(SuccessfulDownloadCount == URLs.Count, "Test Passed: We downloaded the correct number of product images.");
            Assert.IsTrue(failedDownloadList.Count == 0, "Test Passed:  Every single Product Image was downloaded.");
        }

        [TestMethod]
        public void T3_FailToDownloadProductImage() {

            // ----------------------------------------------------------------------
            // Arange
            // ----------------------------------------------------------------------            
            int SuccessfulDownloadCount = 0;
            List<string> failedDownloadList = new List<string>();

            Excel imageLinks = new Excel();

            // Create a fake URL to gaurantee failure during download
            List<string> URLs = new List<string>() { 
                "http://Vehiclepartimages.com/pmdt/DMT/images/01100WH_InvalidURL_UnitTesting.jpg" 
            };

            // ----------------------------------------------------------------------
            // Act
            // ----------------------------------------------------------------------
            SuccessfulDownloadCount = imageLinks.DownloadProductImages(URLs, out failedDownloadList);

            if (failedDownloadList != null && failedDownloadList.Count > 0) {

                // Save the list of files that failed to download
                System.IO.File.WriteAllLines(this.ExceptionList, failedDownloadList);
            }

            // ----------------------------------------------------------------------
            // Assert
            // ----------------------------------------------------------------------
            Assert.IsTrue(SuccessfulDownloadCount == 0, "Test Passed: We failed to download the invalid image URL.");
            Assert.IsTrue(failedDownloadList.Count == 1, "Test Passed:  We counted exactly 1 failed download attempt.");
        }
    }
}
