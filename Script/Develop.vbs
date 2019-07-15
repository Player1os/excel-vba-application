Option Explicit

' Declare local variables.
Dim vProjectDirectoryPath
Dim vBuildConfiguration

' Retrieve the project's directory path.
vProjectDirectoryPath = GetLocalProjectDirectoryPath()

' If the main workbook is already open, notify the user and exit.
If IsMainWorkbookOpen(vProjectDirectoryPath) Then
	Call MsgBox("The main workbook is already open in a different process and must be closed before proceeding.", vbExclamation)
	Call WScript.Quit
End If

' Load the build configuration from the build configuration xml document.
Set vBuildConfiguration = LoadBuildConfiguration(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Build.xml"))

' Create the main workbook.
Call CreateMainWorkbook(vProjectDirectoryPath, vBuildConfiguration)

' Create the execute script.
Call CreateExecuteScript(vProjectDirectoryPath, vBuildConfiguration)

' Set the required environment variables.
With vWScriptShell.Environment("PROCESS")
	' Indicates that the project is to be run in debug mode.
	.Item("APP_IS_DEBUG_MODE_ENABLED") = "TRUE"
	' Indicates that the project is to be run in background mode.
	If vBuildConfiguration("IsBackgroundModeEnabled") Then
		.Item("APP_IS_BACKGROUND_MODE_ENABLED") = "TRUE"
	End If
	' Stores the project name.
	.Item("APP_PROJECT_NAME") = "[Develop] " & vBuildConfiguration("ProjectName")
End With

' Inialize a backup instance of the Excel application for other workbooks to use.
With CreateObject("Excel.Application")
	' Open the project's main workbook in debug mode.
	With CreateObject("Excel.Application")
		' Display the application window.
		Call ShowExcelApplication(.Application)

		' Open the main workbook file the prepared password.
		Call .Workbooks.Open(GetMainWorkbookFilePath(vProjectDirectoryPath), , , , GetMainWorkbookFilePassword(vBuildConfiguration))

		' Wait for the main workbook to be closed.
		Do While .Workbooks.Count > 0
			Call WScript.Sleep(1000)
		Loop
	End With

	' Export the project main workbook's modules.
	Call ExportMainWorkbookModules(vProjectDirectoryPath, vBuildConfiguration)
End With
