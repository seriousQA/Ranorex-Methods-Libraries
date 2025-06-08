/*
 * Created by Ranorex
 * User: seriousQA
 * Date: 03.10.2024
 * Time: 17:34
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using Microsoft.VisualBasic.FileIO;
using System.Management;
using System.ComponentModel;
using DevExpress.Pdf;
	
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;
using Ranorex.Core.Repository;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserMethods
{
    /// <summary> Methods that can be useful in test automation. </summary>
    [UserCodeCollection]
    public class METHODSlib
    {
        #region global_variable
        // There should be the reference to your Ranorex repository
        public static yourProjectRepository repo_global = yourProjectRepository.Instance;
        #endregion

        #region Data
        // There are methods to work with diff. data types. E.g. characters, string, integer.

        /// <summary> Remove the last characters of string. </summary>
		/// <param name="myString"> string, actual string value. </param>
        /// <param name="charCount"> int, amount of characters that have to be removed. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string removeCharsInString(string myString, int charCount)
        {
            string subString = myString.Substring(0, myString.Length - charCount);
            return subString;
        }

        /// <summary> Get first characters of string. </summary>
		/// <param name="myString"> string, actual string value. </param>
        /// <param name="charCount"> int, amount of characters that we need. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getFirstChars(string myString, int charCount)
        {
            string firstChars = myString.Substring(digitsNum);
            return firstChars;
        }

        /// <summary> Get last characters of string. </summary>
		/// <param name="myString"> string, actual string value. </param>
        /// <param name="charCount"> int, amount of characters that we need. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getLastChars(string myString, int charCount)
        {
            string lastChars = myString.Substring(0, myString.Length - digitsNum);
            return lastChars;
        }

        /// <summary> Concatenate two strings. </summary>
		/// <param name="myString1"> string, value #1. </param>
        /// <param name="myString2"> string, value #2. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string concatenateStrings(string myString1, string myString2)
        {
            string conString = myString1 + myString2;
            return conString;
        }

        /// <summary> Compare strings. </summary>
		/// <param name="actual"> actual string. </param>
        /// <param name="expected"> expected string. </param>      
    	[UserCodeMethod]
        public static void compareStrings(string actual, string expected)
        {
            if (actual.Equals(expected))
            {
                Report.Success("Actual value (" + actual + ") is correct.")
            } else {
                Report.Failure("Actual value (" + actual + ") is NOT correct. Expected value is " + expected);
            }
        }

        /// <summary> Validate the attribute value of repoItem isn't null or empty. </summary>
		/// <param name="myRepoItem"> Ranorex repository element. </param>   
        /// <param name="attributeName"> string. </param>
    	[UserCodeMethod]
        public static void validateAttIsNullOrEmpty(Ranorex.Adapter myRepoItem, string attributeName)
        {
            if (String.IsNullOrEmpty(myRepoItem.GetAttributeValue<string>(attributeName)))
            {
                Report.Log(ReportLevel.Error, attributeName + " value of " + myRepoItem.ToString() + " is null or empty.");
            } else {
                Report.Log(ReportLevel.Success, attributeName + " value of " + myRepoItem.ToString() + " isn't null or empty.");
            }
        }

        /// <summary> Get a new string in which all the characters are converted to lowercase. </summary>
        /// <param name="myString"> string, my value. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getStringToLower(string myString)
        {
            string myStringToLower = myString.ToLower();
            return myStringToLower;
        }

        /// <summary> Get a new string in which all the characters are converted to uppercase. </summary>
        /// <param name="myString"> string, my value. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getStringToUpper(string myString)
        {
            string myStringToUpper = myString.ToUpper();
            return myStringToUpper;
        }
        #endregion

        #region DateTime
        // There are methods to work with date, time.

        /// <summary> Get current date with specified format. </summary>
		/// <param name="dateFormat"> string, the date format. E.g. "MM/dd/yyyy". </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getCurrentDate(string dateFormat)
        {
            string currDate = "";
            try {
                System.DateTime date = System.DateTime.Now;
                currDate = date.ToString(dateFormat);
            } catch (Exception inputExp) {
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);
            }
            return currDate;
        }

        /// <summary> Get different date format. </summary>
		/// <param name="myDate"> string, the date in format "dd.MM.yyyy". </param>
        /// <param name="dateFormat"> string, the date format. E.g. "MM'/'yyyy". </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getDateInNewFormat(string myDate, string dateFormat)
        {
            string dateInNewFormat = "";
            try {
                dateInNewFormat = System.DateTime.ParseExact(myDate, "dd.MM.yyyy", System.Globalization.CultureInfo.InvariantCulture).ToString(dateFormat);
            } catch (Exception inputExp) {
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);
            }
            return dateInNewFormat;
        }

        /// <summary> Add months to date. </summary>
        /// <param name="myYear"> int, year in format "YYYY". </param>
        /// <param name="myMonth"> int, month in format "MM". </param>
        /// <param name="myDay"> int, month in format "dd". </param>
		/// <param name="monthsNum"> int, the number of months to add. </param>
        /// <param name="dateFormat"> string, the date format. E.g. "MM'/'yyyy". </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string addMonths(int myYear, int myMonth, int myDay, int monthsNum, string dateFormat)
        {
            string newDate = "";
            try {
                var myDate = new DateTime(myYear, myMonth, myDay);
                newDate = myDate.AddMonths(monthsNum).ToString(dateFormat);
            } catch (Exception inputExp) {
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);
            }
            return newDate;
        }

        /// <summary> Get the time of file creation. </summary>
        /// <param name="filePath"> string, file path. E.g. @"C:\ranorex_testdata\". </param>
		/// <param name="fileName"> string, the name of file. E.g. "myPDFName". </param>
        /// <param name="fileFormat"> string, file format. E.g. "pdf". </param>
        /// <param name="dateFormat"> string, date format. E.g. "dd.MM.yyyy HH:mm:ss". </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getFileCreationTime(string filePath, string fileName, string fileFormat, string dateFormat)
        {
            string currFile = filePath + fileName + "." + fileFormat;
            string currDate = "";
            try {
                System.DateTime actualCreationDate = File.GetCreationTime(currFile);
                System.DateTime formattedCreationDate = Convert.ToDateTime(actualCreationDate);
                currDate = formattedCreationDate.ToString(dateFormat);
                Report.Log(ReportLevel.Info, fileName + " is created at: " + currDate);
            } catch (Exception inputExp) {
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);
            }
            return currDate;
        }
        #endregion        

        #region Files
        // There are methods to work with files, folders.

        /// <summary> Get file[0] name if exists. </summary>
		/// <param name="dirPath"> string, directory path. E.g. @"C:\ranorex_testdata\". </param>  
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
        public static string getFileName(string dirPath)
        {
            try {
                string[] files = Directory.GetFiles(dirPath);
                if (!files.Length.ToString().Equals("0"))
                {
                    string fileName = Path.GetFileName(files[0]).ToString();
                    Report.Log(ReportLevel.Success, "File name  is: " + fileName);
                    return fileName;
                } else {
                    Report.Log(ReportLevel.Success, "Unable to find files in folder.");
                    return null;
                }
            } catch (Exception inputExp) {
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);
            }
        }
        #endregion

        #region DevExpress
        // There are methods to work with DevExpress tool.

        /// <summary> Get a coordinates of text in the PDF file. </summary>
        /// <param name="filePath"> string, file path. E.g. @"C:\ranorex_testdata\". </param>
        /// <param name="fileName"> string, the name of file. E.g. "myPDFName". </param>
        /// <param name="text"> string, the text that we have to find. </param>
    	[UserCodeMethod]
        public static void getCoordinates(string filePath, string fileName, string text)
        {
            string pdfFile = filePath + fileName + ".pdf";

            // specify search parameters
            PdfTextSearchParameters searchParameters = new PdfTextSearchParameters();
            searchParameters.CaseSensitive = true;
            searchParameters.WholeWords = true;

            // declare a list to store the words and their coordinates
            List<Tuple<string, int, int>> wordCoordinatesList = new List<Tuple<string, int, int>>();
            using (PdfDocumentProcessor processor = new PdfDocumentProcessor())
            {
                processor.LoadDocument(pdfFile);
                PdfTextSearchResults currWord = processor.FindText(text, searchParameters);
                for (int i = 0, i < 1; i++)
                {
                    // retrieve the number of the pages where the word is located
                    int pageNumber = currWord.PageNumber;

                    // retrieve the rectangle encompassing the word
                    var elementRectangle = currWord.Rectangles[i];

                    // variables to describe the upper left corner of the word surrounding rectangle
                    int rectTopLeftX = (int)elementRectangle.BoundingRectangle.TopLeft.X;
                    int rectTopLeftY = (int)elementRectangle.BoundingRectangle.TopLeft.Y;

                    // variables to describe the lower right corner of the word surrounding rectangle
                    int rectBottomRightX = (int)elementRectangle.BoundingRectangle.BottomRight.X;
                    int rectBottomRightY = (int)elementRectangle.BoundingRectangle.BottomRight.Y;

                    // add the segment content and its coordinates to the list
                    wordCoordinatesList.Add(new Tuple<string, int, int>(text + "startPosition XY: ", rectTopLeftX, rectTopLeftY));
                    wordCoordinatesList.Add(new Tuple<string, int, int>(text + "endPosition XY: ", rectBottomRightX, rectBottomRightY));

                    Report.Log(ReportLevel.Info, "pageNumber: " + pageNumber);
                    Report.Log(ReportLevel.Info, "startPosition X: " + rectTopLeftX);
                    Report.Log(ReportLevel.Info, "startPosition Y: " + rectTopLeftY);
                    Report.Log(ReportLevel.Info, "endPosition X: " + rectBottomRightX);
                    Report.Log(ReportLevel.Info, "endPosition Y: " + rectBottomRightY);
                }
                processor.CloseDocument();
            }
        }
        #endregion

        #region randomGenerators

        /// <summary> Generate random string. </summary>
        /// <param name="charsCount"> int, amount of characters. </param>
        /// <returns> Return: string. </returns>
        [UserCodeMethod]
        public static string generateRandomString(int charsCount)
        {
            var allowedChars = "ABCDEFGHJKLMNOPRSTUVWXYZ";
            var randomChars = new char[charsCount];
            var random = new Random();
            for (var i = 0; i < charsCount; i++) {
                randomChars[i] = allowedChars[random.Next(0, allowedChars.Length)];
            }
            var builder = new StringBuilder();
            builder.Append(randomChars);
            string result = builder.ToString();
            return result;
        }

        /// <summary> Generate random integer. </summary>
		/// <param name="digitsCount"> int, amount of digits. </param>
        /// <returns> Return: int. </returns>
    	[UserCodeMethod]
        public static int generateRandomInt(int digitsCount)
        {
            var allowedDigits = "0123456789";
            var randomDigits = new char[digitsCount];
            var random = new Random();
            for (var i = 0; i < digitsCount; i++) {
                randomDigits[i] = allowedDigits[random.Next(0, allowedDigits.Length)];
            }
            var builder = new StringBuilder();
            builder.Append(randomDigits);
            string result = builder.ToString();
            return result;
        }
        #endregion

        #region Microsoft.Office.Interop.Excel

        /// <summary> Save elapsed time of method execution into Excel file. </summary>
        [UserCodeMethod]
        public static string saveDataIntoExcelFile()
        {
            var watcher = System.Diagnostics.Stopwatch.StartNew();
            Excel.Application myExcel = new Excel.Application();
            Excel.Workbook myWorkbook = myExcel.Workbooks.Open("C:\\your\\path\\to\\file.xlsx");
            Excel.Worksheet myWorksheet = (Excel.Worksheet)myWorkbook.Worksheets[1];
            Excel.Range myRange = (Excel.Range).myWorksheet.Columns[1];
            Excel.Range lastCell = (Excel.Range).myWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            watcher.Stop();
            var elapsedMs = watcher.ElapsedMilliseconds;
            ((Excel.Range)myWorksheet.Cells[lastCell.Row + 1, 1]).Value = "test";
            ((Excel.Range)myWorksheet.Cells[lastCell.Row + 1, 2]).Value = elapsedMs;

            myWorkbook.Close(true);
            myExcel.Quit();
    	}
        #endregion
    }
}
