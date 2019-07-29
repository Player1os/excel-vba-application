Attribute VB_Name = "Runtime"
Option Explicit
Option Private Module

' Requires Controller.

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

Private Const vDefaultErrorNumber As Long = 10000

Private vIsErrorStored As Boolean
Private vIsErrorIntercepted As Boolean

Private vStoredErrorNumber As Long
Private vStoredErrorSource As String
Private vStoredErrorDescription As String
Private vStoredErrorMessage As String

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

Public Function Username() As String
    Username = WScriptShell().Environment("PROCESS")("USERNAME")
End Function

Public Function ComputerName() As String
    ComputerName = WScriptShell().Environment("PROCESS")("COMPUTERNAME")
End Function

Public Function ConfigFilePath() As String
    ConfigFilePath = FileSystemObject().BuildPath(ThisWorkbook.Path, "Config.xml")
End Function

Public Function ErrorFilePath() As String
    ErrorFilePath = FileSystemObject().BuildPath(ThisWorkbook.Path, "Error.log")
End Function

Public Function IconFilePath() As String
    With FileSystemObject()
        IconFilePath = .BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.ico")
    End With
End Function

Public Function BaseHtmlFilePath() As String
    With FileSystemObject()
        BaseHtmlFilePath = .BuildPath(.BuildPath(ThisWorkbook.Path, "Assets"), "Main.html")
    End With
End Function

Public Sub SetActiveWindowIcon()
    ' Send the api message that loads and sets an icon for the currently active window.
    Call SendMessageA(GetActiveWindow(), &H80, 0, ExtractIconA(0, IconFilePath(), 0))
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

Public Sub SetErrorMessage( _
    ByVal vMessage As String _
)
    vStoredErrorMessage = vMessage
End Sub

Public Sub ClearErrorMessage()
    vStoredErrorMessage = vbNullString
End Sub

Public Sub RaiseError( _
    ByRef vSource As String, _
    ByRef vDescription As String, _
    Optional ByVal vMessage As String = vbNullString _
)
    ' Store the error message.
    If vMessage <> vbNullString Then
        vStoredErrorMessage = vMessage
    End If

    ' Raise the error with the correct number and description.
    Call VBA.Err.Raise(vDefaultErrorNumber, vSource, vDescription)
End Sub

Public Sub StoreError()
    ' Check whether the error had already been intercepted.
    If Not vIsErrorIntercepted Then
        ' Start debugging if in debug mode.
        Debug.Assert Not IsDebugModeEnabled()

        ' Set the error caught flag.
        vIsErrorIntercepted = True
    End If

    ' Set the error stored flag.
    vIsErrorStored = True

    ' Store the current error parameters.
    vStoredErrorNumber = VBA.Err.Number
    vStoredErrorSource = VBA.Err.Source
    vStoredErrorDescription = VBA.Err.Description
End Sub

Public Sub ReRaiseError()
    ' Verify that an error is stored.
    If vIsErrorStored Then
        ' Reset the error stored flag.
        vIsErrorStored = False

        ' ReRaise an error with the stored error parameters.
        Call VBA.Err.Raise(vStoredErrorNumber, vStoredErrorSource, vStoredErrorDescription)
    End If
End Sub

Public Function ParseNavigatePath( _
    ByRef vNavigatePath As String _
) As Dictionary
    ' Declare local variables.
    Dim vPath As String
    Dim vParametersPortionIndex As Long
    Dim vParameterEntry As Variant
    Dim vParsedParameterEntry() As String
    Dim vParameters As New Dictionary

    vPath = vNavigatePath
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

    Set ParseNavigatePath = New Dictionary
    With ParseNavigatePath
        .Item("Path") = vPath
        Set .Item("Parameters") = vParameters
    End With
End Function

Public Function GenerateNavigatePath( _
    ByRef vPath As String, _
    ByRef vParameters As Dictionary _
) As String
    ' Declare local variables.
    Dim vParameterKey As Variant
    Dim vParameterKeyIndex As Long
    Dim vParameterEntries() As String

    GenerateNavigatePath = vPath
    If vParameters.Count > 0 Then
        ReDim vParameterEntries(0 To vParameters.Count - 1)
        vParameterKeyIndex = LBound(vParameterEntries)
        For Each vParameterKey In vParameters.Keys()
            vParameterEntries(vParameterKeyIndex) = vParameterKey & "=" & vParameters(vParameterKey)
            vParameterKeyIndex = vParameterKeyIndex + 1
        Next
        GenerateNavigatePath = GenerateNavigatePath & "?" & Join(vParameterEntries, "&")
    End If
End Function

Public Sub Navigate( _
    ByRef vNavigatePath As String _
)
    ' Declare local variables.
    Dim vHasErrorOccurred As Boolean
    Dim vPath As String
    Dim vParameters As Dictionary

    ' Configure error handling.
    On Error GoTo HandleError:

    ' Extract the query parameters if available.
    With ParseNavigatePath(vNavigatePath)
        vPath = .Item("Path")
        Set vParameters = .Item("Parameters")
    End With

    ' Pass the path and parameters to the user defined controller.
    Call Controller.Navigate(vPath, vParameters)

Terminate:
    ' Reset error handling
    On Error GoTo 0

    ' If an error had occurred and the userform is visible, hide the userform.
    If vHasErrorOccurred Then
        If ThisUserForm.Visible Then
            Call Unload(ThisUserForm)
        End If
    End If

    ' Exit the procedure.
    Exit Sub

HandleError:
    ' Set the error flag.
    vHasErrorOccurred = True

    ' Check whether the error had already been caught.
    If vIsErrorIntercepted Then
        ' Reset the error caught flag.
        vIsErrorIntercepted = False
    Else
        ' Start debugging if in debug mode.
        Debug.Assert Not IsDebugModeEnabled()
    End If

    ' Handle error reporting.
    Call Controller.HandleError(vPath, vParameters, vStoredErrorMessage)

    ' Clear the stored error.
    vIsErrorStored = False
    vStoredErrorNumber = 0
    vStoredErrorSource = vbNullString
    vStoredErrorDescription = vbNullString
    vStoredErrorMessage = vbNullString

    ' Terminate error handling.
    Resume Terminate:
End Sub
