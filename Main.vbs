Option Explicit

' Define constants.
Const xlMaximized = -4137

' Load external modules.
Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")

' Set the environment variable that indicates the app has been executed by the initialization script.
vWScriptShell.Environment("PROCESS")("APP_IS_EXECUTED_BY_SCRIPT") = "TRUE"

' Determine the project directory path and main workbook file name.
Dim vProjectDirectoryPath: vProjectDirectoryPath = vFileSystemObject.GetParentFolderName(WScript.ScriptFullName)
Dim vMainWorkbookFileName: vMainWorkbookFileName = Left( _
	WScript.ScriptName, _
	Len(WScript.ScriptName) - Len(vFileSystemObject.GetExtensionName(WScript.ScriptName)) _
) & "xlsm"

' Initialize a separate instance of the Excel application.
With CreateObject("Excel.Application")
	' Open the main project workbook.
	Call .Workbooks.Open(vProjectDirectoryPath & "\" & vMainWorkbookFileName)

	' Make visible the application window.
	.Visible = True

	' Bring the application window to the foreground.
	Call vWScriptShell.AppActivate(.Caption)

	' Maximize the application window.
	.ActiveWindow.WindowState = xlMaximized
End With
