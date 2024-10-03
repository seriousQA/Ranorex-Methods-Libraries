/*
 * Created by Ranorex
 * User: seriousQA
 * Date: 19.12.2018
 * Time: 12:17
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
using System.IO;
using System.Diagnostics;
using Microsoft.Win32;
using System.Security.Principal;
using Microsoft.VisualBasic.FileIO;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace SETUP
{
    /// <summary> SETUP code collection. </summary>
    [UserCodeCollection]
    public class SETUPlib
    {	
    	/// <summary> Run 32-bit or 64-bit application (AUT). </summary>
    	/// <param name="fullExePathX86"> full path to *.exe file (32-bit application).</param>
		/// <param name="shortBinPathX86"> short path to Bin folder (32-bit application).</param>
		/// <param name="fullExePathX64"> full path to *.exe file (64-bit application).</param>
		/// <param name="shortBinPathX64"> short path to Bin folder (64-bit application).</param>    	
    	[UserCodeMethod]
    	public static void openApp(string fullExePathX86, string shortBinPathX86, string fullExePathX64, string shortBinPathX64)
    	{
    		if (File.Exists(fullExePathX86))
    		{
    	    	Host.Local.RunApplication(fullExePathX86, "", shortBinPathX86, false);
    	    	Delay.Seconds(3);
    		}
    		else
    		{
    	    	Host.Local.RunApplication(fullExePathX64, "", shortBinPathX64, false);
    	    	Delay.Seconds(3);
    		}
    	}    	   		
    	
		/// <summary> Delete from the registry the key for application. </summary>
		/// <param name="regeditFolder"> the name of registry folder of your application. </param>
		[UserCodeMethod]
    	public static void deleteRegKey(string regeditFolder)
    	{
			RegistryKey keyName  = Microsoft.Win86.Registry.CurrentUser.OpenSubKey(@"SOFTWARE", true);
    		if (keyName == null)
    		{
    			; // do nothing
    		}
			else 
			{
				keyName.DeleteSubKeyTree(regeditFolder, false);
				keyName.Close();
			}
    	}
    	
		/// <summary> Remove such traces of a test program like \Documents, \Roaming and s.o. </summary>
		/// <param name="regeditFolder"> the name of your application. </param>
		/// <param name="pathDocuments"> e.g. @"c:\%HOMEPATH%\Documents\". </param>
		/// <param name="pathRoaming"> e.g. @"c:\%HOMEPATH%\AppData\Roaming\". </param>
    	[UserCodeMethod]
    	public static void RMDIR(string regeditFolder, string pathDocuments, string pathRoaming)
    	{    		
    		string pathAppDocuments = pathDocuments + regeditFolder;
    		if (Directory.Exists(Environment.ExpandEnvironmentVariables(pathAppDocuments)))
    		{
    			DirectoryInfo Documents = new DirectoryInfo(Environment.ExpandEnvironmentVariables(pathAppDocuments));
        		Documents.Delete(true);
    		}
			
			string pathAppRoaming = pathRoaming + regeditFolder;
			pathAppRoaming = Environment.ExpandEnvironmentVariables(pathAppRoaming);
			
			if (Directory.Exists(pathAppRoaming))
			{
				// search for *.ini	in folder and delete	
				string[] fileList1 = Directory.GetFiles(pathAppRoaming, "*.ini");
				foreach (string f1 in fileList1)
				{
					File.Delete(f1);
				}
				
				// search for *.xml	in folder and delete	
				string[] fileList2 = Directory.GetFiles(pathAppRoaming, "*.xml");
				foreach (string f2 in fileList2)
				{
					File.Delete(f2);
				}
			}
    	}    
    	
		/// <summary> Start dialog > Create new project (GUI) </summary>
    	[UserCodeMethod]
    	public static void createNewProject()
    	{
    		var repo = SETUPRepository.Instance; 		
    		repo.StartDialog.CreateNewProjectBtn.Click();
            Delay.Seconds(15);
    	}
    	
		/// <summary> Open a specific project. </summary>
		/// <param name="patchProject"> path to the project. E.g. @"C:\Ranorex\...\Projects\" </param>
		/// <param name="nameProject"> the project name. E.g. "myProject.docx" </param>
    	[UserCodeMethod]
    	public static void openProject(string patchProject, string nameProject)
    	{
			var repo = SETUPRepository.Instance;
    	    repo.StartDialog.OpenProject.Click();
            Delay.Milliseconds(10);
            
            // form ['Open Project']
           	repo.OpenProject.PreviousLocations.Click();
            repo.OpenProject.Text41477.TextValue = patchProject;
            Delay.Milliseconds(10);  
            Keyboard.Press("{Return}");
            repo.OpenProject.Text1148.Click();
            repo.OpenProject.Text1148.TextValue = nameProject;
	    	Delay.Milliseconds(10);

	    	// All files (*.*)
	    	repo.OpenProject.ComboBox1348.Click();
	    	repo.List.AllFiles.MoveTo();
	    	repo.List.AllFiles.Click();
	    	repo.OpenProject.OpenBtn.Click();
	    	Delay.Seconds(3);
    	}
    	
		/// <summary> Resize an AUT window to actual monitor resolution. </summary>
    	[UserCodeMethod]
    	public static void resizeWindow()
    	{
			var repo = SETUPRepository.Instance;			
    		Size resolution = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Size;
    		repo.ProjectWindow.Self.Resize(resolution.Width, resolution.Height);
    	}
    	
		/// <summary> Remove from ..\FolderResult all files created during the test run. </summary>
		/// <param name="projectName"> the test project name. </param>
    	[UserCodeMethod]
    	public static void DeleteFilesFromDIR(string projectName)
    	{
    		string FolderResult = @"c:\\Ranorex\\" + projectName + "\\FolderResult\\";
    		if (Directory.Exists(Environment.ExpandEnvironmentVariables(FolderResult)))
    		{
    			DirectoryInfo dirInfo = new DirectoryInfo(Environment.ExpandEnvironmentVariables(FolderResult));
        		foreach (FileInfo file in dirInfo.GetFiles())
				{
    				file.Delete();
				}
    		}
    	}
		
		/// <summary> Send a dump with a comment. </summary>
		/// <param name="appname"> the dump-client name. E.g. "DmpClientName". </param>
    	[UserCodeMethod]
    	public static void checkDump()
    	{
    		var repo = SETUPRepository.Instance;			
			
			// process list
           	Process[] procList = Process.GetProcessesByName(appname); 
			
			// watching each process
            foreach (System.Diagnostics.Process proc in procList) 
			{
				// if the required process has started
				if(proc.ProcessName.Contains(appname))
				{
     				Report.Info("Dump-client exists.");
     				repo.DmpClient.Comment.Click();
     				repo.DmpClient.Comment.TextValue = "The failure was detected by an automated test.";
					repo.DmpClient.Send.Click();
     			}
     			else
     			{
     				Report.Info("Dump-client doesn't exist.");
     			}
    		}
    	}
    	
		/// <summary> Create a results folder for each AUT. </summary>
		/// <param name="projectName"> the test project name.</param>
    	[UserCodeMethod]
    	public static void createFolderResult(string projectName)
    	{
			string FolderResult = "c:\\Ranorex\\" + projectName + "\\FolderResult\\";
           	FileSystem.CreateDirectory(FolderResult);
            Delay.Milliseconds(10);
    	}  
    }
}
