# Ranorex-Methods-Libraries

The libraries contain the most important and useful C# methods for test automation in Ranorex IDE.
<p align="left">
<img src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/csharp/csharp-original.svg" alt="sharp" height="50"/>
<img src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/dotnetcore/dotnetcore-original.svg" alt="net" height="50"/>
<img src="https://github.com/user-attachments/assets/a862e7aa-2cb9-4075-8a13-a0f210b37747" alt="ranorex" height="50"/>
</p>

## __Requirements:__
* Windows
* Ranorex Studio

## __How to use:__
0) download the libraries (*.cs files);
1) add (or link) this *.cs files in your project;
2) use the methods in yours record or code modules;

## __License:__
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## __SETUPlib__ is a SETUP code collection and includes:
- openApp() to run 32-bit or 64-bit application (AUT);
- deleteRegKey() to delete from the registry the key for application;
- rmDir() to remove such traces of an application (AUT) like \Documents, \Roaming and s.o.;
- createNewProject() to start an application (AUT) and create a new project in it (as GUI method example);
- openProject() to open a specific project in application (AUT) (as GUI method example);
- resizeWindow() to resize an AUT window to actual monitor resolution;
- deleteFilesFromDir() to remove from \FolderResult all files created during the test run;
- checkDump() to send a dump with a comment (as GUI method example);
- createFolderResult() to create a results folder for each test product (AUT).

__SetKeyboardLayout__ TestModule is using to change Windows OS keyboard layout using C#. 
For example, U.S. English layout = "00010409", German layout = "00010407", Russian layout = "00010419". 
More Info from [MSDN](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-loadkeyboardlayouta).

__TheFirstRun__ TestModule is using to run 32-bit or 64-bit test application and resize window.

__StartPopupWatcher__ TestModule is using to start a PopupWatcher to handle dialogs that can possibly come up during test run.

## __TEARDOWNlib__ is a TEARDOWN code collection and includes:
- closeApp() to find and kill process of test application (AUT);
- cleanUpResultsFolder() to delete a results folder for each AUT;
- closeProjectWindow() to close the test product (as GUI method example);
- closeAppInstances() to close all instances of AUT;
- killProcessTree() to kill a process tree and all children using taskkill;
- killProcessAndChildren() to kill a process and all children using ManagementObjectSearcher.

## __METHODSlib__ is a collection of useful methods for test automation and includes:
#region Data
- removeCharsInString() to remove the last characters of string;
- getFirstChars() to get first characters of string;
- getLastChars() to get last characters of string;
- concatenateStrings() to concatenate two strings;
- compareStrings() to compare strings;
- validateAttIsNullOrEmpty() to validate the attribute value of repoItem isn't null or empty;
- getStringToLower() to get a new string in which all the characters are converted to lowercase;
- getStringToUpper() to get a new string in which all the characters are converted to uppercase;

#region DateTime
- getCurrentDate() to get current date with specified format;
- getDateInNewFormat() to get different date format;
- addMonths() to add months to date (the same you can do with days or years);
- getFileCreationTime() to get the time of file creation;

#region Files
- getFileName() to get file[0] name if exists;

#region DevExpress
- getCoordinates() to get a coordinates of text in the PDF file.

#region randomGenerators
- generateRandomString() to  generate random string;
- generateRandomInt() to generate random integer;
