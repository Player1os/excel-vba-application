Option Explicit

' Load external modules.
Dim vFileSystemObject: Set vFileSystemObject = CreateObject("Scripting.FileSystemObject")
Dim vWScriptShell: Set vWScriptShell = CreateObject("WScript.Shell")

' Determine the project directory path and the execute script file name.
Dim vProjectDirectoryPath: vProjectDirectoryPath = vFileSystemObject.GetParentFolderName(WScript.ScriptFullName)

' Set the environment variable, that indicates that the project is to be run in development mode.
vWScriptShell.Environment("PROCESS")("APP_IS_DEVELOPMENT_MODE_ENABLED") = "TRUE"

' Run the execute script.
Call vWScriptShell.Run(vFileSystemObject.BuildPath(vProjectDirectoryPath, "Debug.vbs"), 0, False)
