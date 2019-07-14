Option Explicit

' Define the project parameter constants.
Const vIsBackgroundModeEnabled = False
Const vMainWorkbookFilePassword = "ExcelVBAApplication"

' Declare local variables.
Dim vMainWorkbookFilePath

' Determine the path to the project's directory.
With CreateObject("Scripting.FileSystemObject")
	vMainWorkbookFilePath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), "App.xlsm")
End With

' Inialize a backup instance of the Excel application for other workbooks to use.
With CreateObject("Excel.Application")
	' Initialize an isolated instance of the Excel application and open the main workbook within it.
	With CreateObject("Excel.Application")
		' Check whether background mode is enabled.
		If Not vIsBackgroundModeEnabled Then
			' Make the application window visible and bring it to the forefront
			.Visible = True
			Call CreateObject("WScript.Shell").AppActivate(.Caption)
		End If

		' Open the main workbook file in read-only mode with the prepared password.
		Call .Workbooks.Open(vMainWorkbookFilePath, , True, , vMainWorkbookFilePassword)
	End With
End With
