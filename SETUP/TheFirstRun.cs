/*
 * Created by Ranorex
 * User: Vasilenok_e
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
    /// <summary>
	/// First Run of App
	/// </summary>
    [TestModule("4691C569-AFE8-4C6E-AA43-5E4ADCCEC90C", ModuleType.UserCode, 1)]
    public class TheFirstRun : ITestModule
    {
       
        public TheFirstRun()
        {
            // Do not delete - a parameterless constructor is required!
        }

		// full path to *.exe file (32-bit application)
		// for example, value="c:\\Program Files (x86)\\...\\Bin\\myApp.exe">
		string _FullExePathX86 = "";
		[TestVariable("8899ccf1-797a-4bfb-ad4f-82495c5923ed")]
		public string FullExePathX86
		{
			get { return _FullExePathX86; }
			set { _FullExePathX86 = value; }
		}

		// short path to Bin folder (32-bit application)
		// for example, value="c:\\Program Files (x86)\\...\\Bin">
		string _ShortBinPathX86 = "";
		[TestVariable("8908a0af-e0cf-48a8-9634-0cecac32a025")]
		public string ShortBinPathX86
		{
			get { return _ShortBinPathX86; }
			set { _ShortBinPathX86 = value; }
		}

		// full path to *.exe file (64-bit application)
		// for example, value="c:\\Program Files\\...\\Bin\\myApp.exe">
		string _FullExePathX64 = "";
		[TestVariable("96a7340c-31ee-4a84-90c9-27ec75fec6e2")]
		public string FullExePathX64
		{
			get { return _FullExePathX64; }
			set { _FullExePathX64 = value; }
		}

		// short path to Bin folder (64-bit application)
		// for example, value="c:\\Program Files\\...\\Bin">
		string _ShortBinPathX64 = "";
		[TestVariable("aa58fc42-a145-462c-8f43-d831d3470501")]
		public string ShortBinPathX64
		{
			get { return _ShortBinPathX64; }
			set { _ShortBinPathX64 = value; }
		}

		// if the product and process names do not match, then it is the product name
		// for example, value="myApp"
		string _product = "";
		[TestVariable("dab256ea-95b3-4a1d-8b04-ce4720bfbe1e")]
		public string product
		{
			get { return _product; }
			set { _product = value; }
		}

        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            var repo = SETUPRepository.Instance;
            
            // run application
            SETUPlib.FirstOpenApp(FullExePathX86, ShortBinPathX86, FullExePathX64, ShortBinPathX64);
            Delay.Seconds(5);
            
            // set path to Templates
            SETUPlib.FirstRunDialogPath(product);
            Delay.Seconds(5);
            
			// create new project
            SETUPlib.StartDialogCreateNewProject();
            
            // resize window
            SETUPlib.ResizeWindow();
            Delay.Seconds(1);

        }
    }
}
