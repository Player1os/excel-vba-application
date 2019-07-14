Option Explicit

' Load external modules.
Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim vWscriptShell: Set vWscriptShell = CreateObject("WScript.Shell")

' Define a collection of folder paths.
Dim vLogFolderPaths: Set vLogFolderPaths = CreateObject("Scripting.Dictionary")

' Load the runtime parameters.
Dim vTaskLogDirectoryPath: vTaskLogDirectoryPath = _
	vWscriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH%")
Dim vMaximumTaskLogCount: vMaximumTaskLogCount = _
	CLng(vWscriptShell.ExpandEnvironmentStrings("%APP_TASK_RUNNER_MAXIMUM_TASK_LOG_COUNT%"))

' Store the current subdirectories of the specified task's log directory.
Dim vLogFolder
Dim vLogFolderIndex
vLogFolderIndex = 1
For Each vLogFolder In vFileSystemObject.GetFolder(vTaskLogDirectoryPath).SubFolders
	Call vLogFolderPaths.Add(vLogFolderIndex, vLogFolder.Path)
	vLogFolderIndex = vLogFolderIndex + 1
Next

' Remove the oldest extraneous log subdirectories of the specified task's log directory.
For vLogFolderIndex = 1 To vLogFolderPaths.Count - vMaximumTaskLogCount
	Call vFileSystemObject.DeleteFolder(vLogFolderPaths(vLogFolderIndex))
Next

