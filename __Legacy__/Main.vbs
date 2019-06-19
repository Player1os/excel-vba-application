Option Explicit

' Define constants.
Const vVbCritical = 16

' Store the project's directory path and message box title.
Dim vMessageBoxTitle: vMessageBoxTitle = "[" & WScript.ScriptName & "] " & vProjectDirectoryPath

' Check for the existance of the password file.
Dim vPasswordFilePath: vPasswordFilePath = vProjectDirectoryPath & "\Password.txt"
If vFileSystemObject.FileExists(vPasswordFilePath) Then
	' Load the password file contents into the project password environment variable.
	With CreateObject("ADODB.Stream")
		.CharSet = "UTF-8"
		Call .Open
		Call .LoadFromFile(vPasswordFilePath)
		vWScriptShell.Environment("PROCESS")("APP_PROJECT_PASSWORD") = .ReadText()
		Call .Close
	End With
ElseIf vWScriptShell.Environment("PROCESS")("APP_IS_DEVELOP_MODE_ENABLED") = "TRUE" Then
	' If develop mode is enabled, a password file is required.
	Call MsgBox("The '" & vPasswordFilePath & "' file does not exist.", vVbCritical, vMessageBoxTitle)
	Call WScript.Quit(1)
End If

' TODO: Implement additional logic.

' Set additional environment variables.
With vWScriptShell.Environment("PROCESS")
	' TODO: Set environment variables to be passed to the main project workbook.
End With
