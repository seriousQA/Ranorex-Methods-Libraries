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
        // Reference to the Ranorex repository
        var repo = myRepository.Instance;
        #endregion

        #region Data
        // There are methods to work with diff. data types. E.g. characters, string, integer.

        /// <summary> Remove the last characters of string. </summary>
		/// <param name="myString"> string, actual string value. </param>
        /// <param name="charCount"> int, amount of characters that have to be removed. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
    	public static string removeCharsInString (string myString, int charCount)
    	{
			string subString = myString.Substring(0, myString.Length - charCount);
            return subString;
    	}

        /// <summary> Concatenate two strings. </summary>
		/// <param name="myString1"> string, value #1. </param>
        /// <param name="myString2"> string, value #2. </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
    	public static string concatenateStrings (string myString1,string myString2)
    	{
			string conString = myString1 + myString2;
            return conString;
    	}
        #endregion

        #region DateTime
        // There are methods to work with date, time.

        /// <summary> Get current date with specified format. </summary>
		/// <param name="dateFormat"> string, the date format. E.g. "MM/dd/yyyy". </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
    	public static string getCurrentDate (string dateFormat)
    	{
			string currDate = "";
            try{
                System.DateTime date = System.DateTime.Now;
                currDate = date.ToString(dateFormat);
            }catch(Exception inputExp){
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);                
            }
            return currDate;
    	}

        /// <summary> Get the time of file creation. </summary>
        /// <param name="filePath"> string, file path. E.g. @"C:\ranorex_testdata\". </param>
		/// <param name="fileName"> string, the name of file. E.g. "myPDFName". </param>
        /// <param name="fileFormat"> string, file format. E.g. "pdf". </param>
        /// <param name="dateFormat"> string, date format. E.g. "dd.MM.yyyy HH:mm:ss". </param>
        /// <returns> Return: string. </returns>
    	[UserCodeMethod]
    	public static string getFileCreationTime (string filePath, string fileName, string fileFormat, string dateFormat)
    	{
			string currFile = filePath + fileName + "." + fileFormat;
            string currDate = "";
            try{
                System.DateTime actualCreationDate = File.GetCreationTime(currFile);
                System.DateTime formattedCreationDate = Convert.ToDateTime(actualCreationDate);
                currDate = formattedCreationDate.ToString(dateFormat);
                Report.Log(ReportLevel.Info, fileName + " is created at: " + currDate);
            }catch(Exception inputExp){
                Report.Failure("Stacktrace: " + inputExp.Stacktrace);                
            }
            return currDate;
    	}
        #endregion        

        #region Files
        // There are methods to work with files, folders.

        /// <summary> . </summary>
		/// <param name="myParam"> string. </param>        
    	[UserCodeMethod]
    	public static void doSmth (string myParam)
    	{
			
    	}
        #endregion
    }
}
