Option Explicit

' Define constants.
Const vVbCritical = 16

' Store the project's directory path.
Dim vProjectDirectoryPath: vProjectDirectoryPath = _
	vFileSystemObject.GetParentFolderName(vFileSystemObject.GetParentFolderName(WScript.ScriptFullName))
Dim vMessageBoxTitle: vMessageBoxTitle = "[" & WScript.ScriptName & "] " & vProjectDirectoryPath

' Find the main vbscript file path.
Dim vSelectedFileObject: Set vSelectedFileObject = Nothing
Dim vFileObject
For Each vFileObject In vFileSystemObject.GetFolder(vProjectDirectoryPath).Files
	If Right(vFileObject.Name, 4) = ".vbs" Then
		Exit For
	End If
Next

If vSelectedFileObject Is Nothing Then

End If


' Enable the develop mode environment variable.
vWScriptShell.Environment("PROCESS")("APP_IS_DEVELOP_MODE_ENABLED") = "TRUE"

' Run the main project workbook.
With CreateObject("Excel.Application")
	.Visible = True
	Call .Workbooks.Open(vProjectDirectoryPath & "\Main.xlsm")
End With
