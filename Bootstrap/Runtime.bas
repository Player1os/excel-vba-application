Attribute VB_Name = "Runtime"
Option Explicit
Option Private Module

' Requires ModController.

Private Declare Function GetActiveWindow Lib "user32" () As Integer

Private Declare Function ExtractIconA Lib "shell32.dll" ( _
    ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long _
) As Long

Private Declare Function SendMessageA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function GetWindowLongA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long _
) As Long

Private Declare Function SetWindowLongA Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long

Private Declare Function ShowWindow Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal nCmdShow As Long _
) As Long

Private Const vGwlStyle As Long = -16
Private Const vWsMaximizeBox As Long = &H10000
Private Const vWsMinimizeBox As Long = &H20000
Private Const vWsThickFrame As Long = &H40000
Private Const vWsSystemMenu As Long = &H80000
Private Const vSwShowMaximized As Long = 3

Private vFileSystemObject As FileSystemObject
Private vWScriptShell As Object

Public Function FileSystemObject() As FileSystemObject
    ' Initialize the file system object for use across the project, if needed.
    If vFileSystemObject Is Nothing Then
        Set vFileSystemObject = New FileSystemObject
    End If
    Set FileSystemObject = vFileSystemObject
End Function

Public Function WScriptShell() As Object
    ' Initialize the wscript shell object for use across the project, if needed.
    If vWScriptShell Is Nothing Then
        Set vWScriptShell = CreateObject("WScript.Shell")
    End If
    Set WScriptShell = vWScriptShell
End Function

Public Function IsDebugModeEnabled() As Boolean
    IsDebugModeEnabled = WScriptShell().Environment("PROCESS")("APP_IS_DEBUG_MODE_ENABLED") = "TRUE"
End Function

Public Function IsDeployDebugModeEnabled() As Boolean
    IsDeployDebugModeEnabled = WScriptShell().Environment("PROCESS")("APP_IS_DEPLOY_DEBUG_MODE_ENABLED") = "TRUE"
End Function

Public Function IsBackgroundModeEnabled() As Boolean
    IsBackgroundModeEnabled = WScriptShell().Environment("PROCESS")("APP_IS_BACKGROUND_MODE_ENABLED") = "TRUE"
End Function

Public Function ProjectName() As String
    ProjectName = WScriptShell().Environment("PROCESS")("APP_PROJECT_NAME")
End Function

Public Function NavigatePath() As String
    NavigatePath = WScriptShell().Environment("PROCESS")("APP_NAVIGATE_PATH")
End Function

Public Sub SetActiveWindowIcon()
    ' Declare local variables.
    Dim vIconFilePath As String

    ' Determine the icon file path.
    With FileSystemObject()
        vIconFilePath = .BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.ico")
    End With

    ' Send the api message that loads and sets an icon for the currently active window.
    Call SendMessageA(GetActiveWindow(), &H80, 0, ExtractIconA(0, vIconFilePath, 0))
End Sub

Public Sub PopulateActiveWindowTitlebar()
    ' Declare local variables.
    Dim vFormHandle As Long
    Dim vWindowStyle As Long

    ' Retrieve the form handle of the currently active window.
    vFormHandle = GetActiveWindow()

    ' Retrieve the new window style information for the currently active window.
    vWindowStyle = GetWindowLongA(vFormHandle, vGwlStyle)

    ' Add the desired properties to the retrieved new window style information.
    vWindowStyle = vWindowStyle Or vWsMaximizeBox
    vWindowStyle = vWindowStyle Or vWsMinimizeBox
    vWindowStyle = vWindowStyle Or vWsThickFrame
    vWindowStyle = vWindowStyle Or vWsSystemMenu

    ' Set the configured new window style information to the currently active window.
    Call SetWindowLongA(vFormHandle, vGwlStyle, vWindowStyle)
End Sub

Public Sub MaximizeActiveWindow()
    Call ShowWindow(GetActiveWindow(), vSwShowMaximized)
End Sub

Public Sub Navigate( _
    ByRef vNavigatePath As String _
)
    ' Declare local variables.
    Dim vPath As String
    Dim vParametersPortionIndex As Long
    Dim vParameterEntry As Variant
    Dim vParsedParameterEntry() As String
    Dim vParameters As New Dictionary

    ' Define the navigate path as the first version of the path.
    vPath = vNavigatePath

    ' Extract the query parameters if available.
    vParametersPortionIndex = InStr(vPath, "?")
    If vParametersPortionIndex <> 0 Then
        For Each vParameterEntry In Split(Mid(vPath, vParametersPortionIndex + 1), "&")
            vParsedParameterEntry = Split(vParameterEntry, "=")
            If UBound(vParsedParameterEntry) = 0 Then
                ReDim Preserve vParsedParameterEntry(0 To 1)
                vParsedParameterEntry(1) = vbNullString
            End If
            Call vParameters.Add(vParsedParameterEntry(0), vParsedParameterEntry(1))
        Next
        vPath = Left(vPath, vParametersPortionIndex - 1)
    End If

    ' Pass the path and parameters to the user defined controller.
    Call ModController.Navigate(vPath, vParameters)
End Sub
