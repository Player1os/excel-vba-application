Option Explicit

' Load external Modules.
Dim vOutlookApplication: Set vOutlookApplication = CreateObject("Outlook.Application")
Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")

' Define constants.
Const vOlMailItem = 0
Const vVbCritical = 16

' Disable error handling.
On Error Resume Next

' Allocate an email item.
Dim vMailItem: Set vMailItem = vOutlookApplication.CreateItem(vOlMailItem)
Call vOutlookApplication.Session.Logon

' Prepare the email and send it.
vMailItem.To = vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_ERROR_MAIL_RECIPIENT%")
vMailItem.Subject = "[Task Runner] Failed to execute task '" & vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_NAME%") & "'"
vMailItem.HTMLBody = "<div style=""font-family: Arial; font-size: 10pt"">" _
	& "<p><b>Error message:</b> <code style=""background-color: #eee; color: #c00"">" _
	& WScript.Arguments(0) _
	& "</code></p>" _
	& "<p><b>Task name:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_NAME%") _
	& "</code></p>" _
	& "<p><b>User name:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%USERNAME%") _
	& "</code></p>" _
	& "<p><b>Machine name:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%COMPUTERNAME%") _
	& "</code></p>" _
	& "<p><b>Script file path:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_SCRIPT_FILE_PATH%") _
	& "</code></p>" _
	& "<p><b>Start timestamp:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_START_TIMESTAMP%") _
	& "</code></p>" _
	& "<p><b>End timestamp:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_END_TIMESTAMP%") _
	& "</code></p>" _
	& "<p><b>Return code:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_RETURN_CODE%") _
	& "</code></p>" _
	& "<p><b>Log directory:</b> <code style=""background-color: #eee; color: #c00"">" _
	& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%") _
	& "</code></p>" _
	& "</div>"
Call vMailItem.Send

' Disconnect and disable the imported libraries.
Call vOutlookApplication.Session.Logoff

' Check whether reporting the error was finished without error, otherwise display a message box to the user.
If Err.Number <> 0 Then
	Call MsgBox("An unexpected error had occured while attemtping to report an error regarding task '" _
		& vWScriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_NAME%") & "'", vVbCritical, "Task runner")
End If

' Reenable error handling.
On Error Goto 0
