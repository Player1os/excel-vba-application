Option Explicit

' Declare local variables.
Dim vProjectDirectoryPath
Dim vProjectConfiguration
Dim vDeployDirectoryPath
Dim vAssetDirectoryPath

' Retrieve the project's directory path.
vProjectDirectoryPath = GetProjectDirectoryPath()

' Load the project's configuration from the config workbook.
Set vProjectConfiguration = LoadProjectConfiguration()

' Build the project's main workbook.
Call BuildMainWorkbook(vProjectConfiguration)

With vFileSystemObject


	' Copy the main workbook file to the deploy directory and set as read-only.
	Call .CopyFile(.BuildPath(vProjectDirectoryPath, "App.xlsm"), vProjectConfiguration("DeployDirectoryPath"), True)
	Call .GetFile(.BuildPath("DeployDirectoryPath"))

	' Copy the main workbook file and execute script file to the deploy directory.
	Call .CopyFile(.BuildPath(vProjectDirectoryPath, "Execute.vbs"), vProjectConfiguration("DeployDirectoryPath"), True)

	' Copy the asset directory to the deploy directory, if it exists.
	vAssetDirectoryPath = .BuildPath(vProjectDirectoryPath, "Asset")
	If .FolderExists(vAssetDirectoryPath) Then
		Call .CopyFolder(vAssetDirectoryPath, vProjectConfiguration("DeployDirectoryPath"), True)
	End If
End With
