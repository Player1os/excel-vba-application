Option Explicit

' Load external modules.
Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")

' Determine the project directory path and the execute script file name.
Dim vProjectDirectoryPath: vProjectDirectoryPath = vFileSystemObject.GetParentFolderName(WScript.ScriptFullName)

' Initialize an isolated instance of the Excel application.
With CreateObject("Excel.Application")
	' Open the main project workbook.
	Call .Workbooks.Open(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Config.xlsx"))

	' Make visible the application window.
	.Visible = True

	' Maximize the application window.
	.ActiveWindow.WindowState = xlMaximized

	' Bring the application window to the foreground.
	Call vWScriptShell.AppActivate(.Caption)
End With
