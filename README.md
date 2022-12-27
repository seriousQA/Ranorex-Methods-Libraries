# Ranorex-Methods-Libraries

The libraries contain the most important and useful C# methods for creating autotests in Ranorex IDE.

## __SETUPlib__ is a SETUP code collection and includes:
- FirstOpenApp() to run 32-bit or 64-bit application;
- REGdelete() to delete from the registry the key for application;
- RMDIR() to remove such traces of a test program like \Documents, \Roaming and s.o.;
- StartDialogCreateNewProject() to start a test product and create a new project in it (as GUI method example);
- OpenProjectStartDialog() to open a specific project in test product (as GUI method example);
- ResizeWindow() to resize a window to actually monitor resolution;
- DeleteFilesFromDIR() to remove from \FolderResult all files created during the test run;
- CheckDamp() to send a dump with a comment (as GUI method example);
- CreateFolderResult() to create a ResultFolder for each test product.

__SetKeyboardLayout__ TestModule is using to change Windows OS keyboard layout using C#. 
For example, U.S. English layout = "00010409", German layout = "00010407", Russian layout = "00010419". 
More Info from [MSDN](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-loadkeyboardlayouta).

__TheFirstRun__ TestModule is using to run 32-bit or 64-bit test application and resize window.

