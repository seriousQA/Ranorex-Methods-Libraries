/*
 * Created by Ranorex
 * User: Vasilenok_e
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
using Microsoft.VisualBasic.FileIO;
using System.Management;
using System.ComponentModel;
	
using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace TEARDOWN
{
    /// <summary>
    /// TEARDOWN code collection.
    /// </summary>
    [UserCodeCollection]
    public class TEARDOWNlib
    {
    	
    	/// <summary>
    	/// Find and kill process of test product.
		/// <param name="processName">the name of process without ".exe".</param>
		/// for example, value="myApp">
    	/// </summary>
    	[UserCodeMethod]
    	public static void CloseApp(string processName)
    	{
			// process list
            Process[] procList = Process.GetProcessesByName(processName);
			
			// watching each process 
            foreach (System.Diagnostics.Process anti in procList) 
			{
				// if the required process has started
				if(anti.ProcessName.Contains(processName))
				{
					Report.Info(processName + " exists. Close.");
					anti.Kill();
					Delay.Seconds(1);					
				}
			
				else
				{
            		Report.Info(processName + " doesn't exist. No entity to close.");            
        		}
				Delay.Seconds(2);
        	}
    	}
    	    	
    	/// <summary>
    	/// Delete a ResultFolder for each test product.
		/// <param name="product">the test product name.</param>
    	/// </summary>
    	[UserCodeMethod]
    	public static void CleanupFolderResult(string product)
    	{
    		DirectoryInfo FolderName = new DirectoryInfo("c:\\Ranorex\\" + product + "\\FolderResult\\");    		
    		FolderName.Delete(true);
    	}    	
    	
		/// <summary>
    	/// Close the test product (GUI example).
		/// <param name="product">the test product name.</param>
    	/// </summary>
    	[UserCodeMethod]
    	public static void CloseProjectWindow(string product)
    	{
    		for (int i=0; i<5; i++)
            	{
    				var repo = TEARDOWNRepository.Instance;
					
    				if(repo.ProjectWindow.CloseBtnInfo.Exists())
    					{    				
							Delay.Seconds(1);
							repo.ProjectWindow.CloseBtn.Click();
							Delay.Seconds(4);
    					}
    				else
    					{
    						Report.Info(product + " doesn't exist. No entity to close.");   
    					}
    			}
    	}
    	    	
    	/// <summary>
    	/// Close all instances of application.
		/// <param name="processName">the name of process without ".exe".</param>
		/// for example, value="myApp">
    	/// </summary>
    	[UserCodeMethod]
    	public static void CloseInstancesApp(string processName)
    	{    		
    		if (processName == null)
    		{
    			return;
    		}

    		Process[] processes = Process.GetProcesses();
    			
    		foreach (Process proc in processes)
			{
    			if (proc.ProcessName == processName)
    			{
    				try
    				{
    					Report.Info(proc.ProcessName.ToString(), proc.Id.ToString());
    					proc.Kill();
    					proc.WaitForExit(1000);
    				}
    				catch (ArgumentException)
    				{
    					// process is already exited
    				}
    			}
    		}
			
    	}
    	
    	/// <summary>
    	/// Kill a process tree and all children using taskkill.
    	/// </summary>
		/// <param name="processName">the name of process without ".exe".</param>
		/// for example, value="myApp">
		/// </summary>
    	[UserCodeMethod]
    	public static void KillProcessTree(string processName)
    	{
    		Process.Start(new ProcessStartInfo
						{
						FileName = "taskkill",
						Arguments = $" /im {processName} /f /t"
						}).WaitForExit();
    	}
    	
    	/// <summary>
    	/// Kill a process and all children.
    	/// </summary>
		/// <param name="processName">the name of process without ".exe".</param>
		/// for example, value="myApp">
		/// </summary>    	
    	[UserCodeMethod]
    	public static void KillProcessAndChildren(string processName)
    	{
    		Process[] p = Process.GetProcessesByName(processName);
    		foreach(var proc1 in p)
    		{
				// process ID
    			int pid = proc1.Id;
    		
				// cannot close \system idle process'
				if (pid == 0)
				{
					return;
				}
    		
				ManagementObjectSearcher searcher = new ManagementObjectSearcher
				("Select * From Win32_Process Where ParentProcessID=" + pid);
				ManagementObjectCollection moc = searcher.Get();
				foreach (ManagementObject mo in moc)
				{
					KillProcessAndChildren(mo["ProcessID"].ToString());
				}
				try
				{
					Process proc2 = Process.GetProcessById(pid);
					Report.Info(proc2.ProcessName.ToString(), proc2.Id.ToString());
					proc2.Kill();
				}
				catch (ArgumentException)
				{
					// Process is already exited
				}
    		}
    	}
    }
}
