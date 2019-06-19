Option Explicit

' Load external modules.
Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")

' Define external module constants.
Const xlMaximized = -4137

' Manipulate the environment variables.
With vWScriptShell.Environment("PROCESS")
	' Set environment variable, which indicates that the main workbook has been opened by the execute script.
	.Item("APP_IS_OPENED_BY_EXECUTE_SCRIPT") = "TRUE"

	' Determine whether development mode has been enabled.
	Dim vIsDevelopmentModeEnabled: vIsDevelopmentModeEnabled = .Item("APP_IS_DEVELOPMENT_MODE_ENABLED") = "TRUE"
End With

' Determine the project directory path.
Dim vProjectDirectoryPath: vProjectDirectoryPath = vFileSystemObject.GetParentFolderName(WScript.ScriptFullName)

' Determine the main workbook file path.
Dim vFileObject
Dim vMainWorkbookFilePath
For Each vFileObject In vFileSystemObject.GetFolder(vProjectDirectoryPath).Files
	If vFileSystemObject.GetExtensionName(vFileObject.Name) = "xlsm" Then
		vMainWorkbookFilePath = vFileObject.Path
		Exit For
	End If
Next

' Initialize an isolated instance of the Excel application.
With CreateObject("Excel.Application")
	' Check whether the application should be displayed.
	If vIsDevelopmentModeEnabled Then
		' Make visible the application window.
		.Visible = True

		' Bring the application window to the foreground.
		Call vWScriptShell.AppActivate(.Caption)
	End If

	' Open the main project workbook.
	Call .Workbooks.Open(vMainWorkbookFilePath, , Not vIsDevelopmentModeEnabled)
End With
