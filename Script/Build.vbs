Option Explicit

' Declare local variables.
Dim vProjectDirectoryPath
Dim vBuildConfiguration

' Retrieve the project's directory path.
vProjectDirectoryPath = GetLocalProjectDirectoryPath()

' If the main workbook is already open, notify the user and exit.
If IsMainWorkbookOpen(vProjectDirectoryPath) Then
	Call MsgBox("The main workbook is already open in a different process and must be closed before proceeding.", vbExclamation)
	Call WScript.Quit()
End If

' Load the build configuration from the build configuration xml document.
Set vBuildConfiguration = LoadBuildConfiguration(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Build.xml"))

' Create the main workbook.
Call CreateMainWorkbook(vProjectDirectoryPath, vBuildConfiguration)

' Create the execute script.
Call CreateExecuteScript(vProjectDirectoryPath, vBuildConfiguration)
