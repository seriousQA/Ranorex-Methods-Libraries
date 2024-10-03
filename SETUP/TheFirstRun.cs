/*
 * Created by Ranorex
 * User: seriousQA
 * Date: 05.02.2019
 * Time: 12:44
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
using System.Diagnostics;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace SETUP
{
    /// <summary> Run of app </summary>
    [TestModule("4691C569-AFE8-4C6E-AA43-5E4ADCCEC90C", ModuleType.UserCode, 1)]
    public class TheFirstRun : ITestModule
    {
       
        public TheFirstRun()
        {
            // Do not delete - a parameterless constructor is required!
        }

		// full path to *.exe file (32-bit application)
		// e.g. value="c:\\Program Files (x86)\\...\\Bin\\myApp.exe">
		string _fullExePathX86 = "";
		[TestVariable("8899ccf1-797a-4bfb-ad4f-82495c5923ed")]
		public string fullExePathX86
		{
			get { return _fullExePathX86; }
			set { _fullExePathX86 = value; }
		}

		// short path to Bin folder (32-bit application)
		// for example, value="c:\\Program Files (x86)\\...\\Bin">
		string _shortBinPathX86 = "";
		[TestVariable("8908a0af-e0cf-48a8-9634-0cecac32a025")]
		public string shortBinPathX86
		{
			get { return _shortBinPathX86; }
			set { _shortBinPathX86 = value; }
		}

		// full path to *.exe file (64-bit application)
		// for example, value="c:\\Program Files\\...\\Bin\\myApp.exe">
		string _fullExePathX64 = "";
		[TestVariable("96a7340c-31ee-4a84-90c9-27ec75fec6e2")]
		public string fullExePathX64
		{
			get { return _fullExePathX64; }
			set { _fullExePathX64 = value; }
		}

		// short path to Bin folder (64-bit application)
		// for example, value="c:\\Program Files\\...\\Bin">
		string _shortBinPathX64 = "";
		[TestVariable("aa58fc42-a145-462c-8f43-d831d3470501")]
		public string shortBinPathX64
		{
			get { return _shortBinPathX64; }
			set { _shortBinPathX64 = value; }
		}

		// if the projectName and process names do not match, then it is the projectName name
		// for example, value="myApp"
		string _projectName = "";
		[TestVariable("dab256ea-95b3-4a1d-8b04-ce4720bfbe1e")]
		public string projectName
		{
			get { return _projectNameName; }
			set { _projectNameName = value; }
		}

        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            var repo = SETUPRepository.Instance;
            
            // run application
            SETUPlib.openApp(fullExePathX86, shortBinPathX86, fullExePathX64, shortBinPathX64);
            Delay.Seconds(5);
            
            // create results folder
            SETUPlib.createFolderResult(projectName);
            Delay.Seconds(5);
            
			// create new project
            SETUPlib.createNewProject();
            
            // resize window
            SETUPlib.resizeWindow();
            Delay.Seconds(1);

        }
    }
}
