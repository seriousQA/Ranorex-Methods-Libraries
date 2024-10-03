/*
 * Created by Ranorex
 * User: seriousQA
 * Date: 17.08.2018
 * Time: 18:07
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
using System.Runtime.InteropServices;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace SETUP
{
	/// <summary> Change Windows OS keyboard layout using C#. </summary>
	/// U.S. English layout = "00010409"
	/// German layout = "00010407"
	/// Russian layout = "00010419"
	/// more Info from MSDN (LoadKeyboardLayoutA function (winuser.h))
    [TestModule("7979E921-E8D7-4514-8AD3-2F500B14490D", ModuleType.UserCode, 1)]
    public class SetKeyboardLayout : ITestModule
    {
		// Adding WinAPI functions from user32.dll:
		
        // Retrieves a handle to the foreground window (the window with which the user is currently working). 
        [DllImport("user32.dll")]
		public static extern IntPtr GetForegroundWindow();
		
		// Post a message to the window with the focus.
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

		// Loads a new input locale identifier (formerly called the keyboard layout) into the system.
		[DllImport("user32.dll")]
		static extern int LoadKeyboardLayout(string pwszKLID, uint Flags);
       
        public SetKeyboardLayout()
        {
            // Do not delete - a parameterless constructor is required!
        }

        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
			// English layout name.
            string lang = "00000409";

			int ret = LoadKeyboardLayout(lang, 1);
			
			// WM_INPUTLANGCHANGEREQUEST has a code 0x50
			PostMessage(GetForegroundWindow(), 0x50, 1, ret);
        }
    }
}
