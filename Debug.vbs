Option Explicit

' Load external modules.
Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")

' Determine the project directory path and the execute script file name.
Dim vProjectDirectoryPath: vProjectDirectoryPath = vFileSystemObject.GetParentFolderName(WScript.ScriptFullName)

' Determine the main workbook file name.
Dim vFileObject
Dim vMainWorkbookFileName
For Each vFileObject In vFileSystemObject.GetFolder(vProjectDirectoryPath).Files
	If vFileSystemObject.GetExtensionName(vFileObject.Name) = "xlsm" Then
		vMainWorkbookFileName = vFileObject.Name
		Exit For
	End If
Next

' Set the environment variable, that indicates that the project is to be run in development mode.
vWScriptShell.Environment("PROCESS")("APP_DEBUG_PROJECT_PASSWORD") = "^" & vMainWorkbookFileName & "$"

' Run the execute script.
Call vWScriptShell.Run(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Execute.vbs"), 0, False)
