Option Explicit
Option Private Module

' Requires CUtilityTable
' Requires MRuntimeParameters
' Requires MUtilityFile
' Requires MUtilityMail
' Requires ThisWorkbook

Public Enum EEventHandling
    vNone = 0
    vWorkbookOpen = 1
    vWorkbookBeforeClose = 2
    vWorksheetChange = 3
    vWorksheetSelectionChange = 4
End Enum

Private Enum EDebugMode
    vUnknown = 0
    vEnabled = 1
    vDisabled = 2
    vIncorrectPassword = 3
End Enum

Private Const vApplicationDefinedErrorNumber As Long = 10000

Private vDebugMode As EDebugMode
Private vIsErrorStored As Boolean
Private vStoredErrorNumber As Long
Private vStoredErrorSource As String
Private vStoredErrorDescription As String
Private vStoredErrorMessage As String

Private vIsActive As Boolean
Private vIsInitialized As Boolean
Private vIsQuitRequested As Boolean

Private vReportErrorEmailAddress As String
Private vDeployLocationPath As String
Private vUsername As String
Private vComputerName As String

Private Sub pActionDeploy()
    ' Declare local variables.
    Dim vApplicationPath As String

    ' Execute optional copy instructions.
    Call MRuntimeParameters.Deploy

    ' Copy the application and set to read only.
    vApplicationPath = vDeployLocationPath & "\" & ThisWorkbook.Name
    If VBA.Dir(vApplicationPath) <> VBA.vbNullString Then
        Call VBA.SetAttr(vApplicationPath, vbNormal)
    End If
    Call ThisWorkbook.SaveCopyAs(vApplicationPath)
    Call VBA.SetAttr(vApplicationPath, vbReadOnly)
End Sub

Private Sub pActionExportProject()
    ' Declare local variables.
    Dim vComponent As VBIDE.VBComponent
    Dim vSuffix As String
    Dim vComponentPath As String
    Dim vDirectoryPath As String
    Dim vFileNameValue As Variant

    ' Make sure the project is unprotected.
    If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
        Call VBA.MsgBox("The VBA Project must be unprotected.", vbExclamation)
        Exit Sub
    End If

    ' Remove all previous export files.
    vDirectoryPath = ThisWorkbook.Path & "\Export"
    For Each vFileNameValue In MUtilityFile.DirectoryContents(vDirectoryPath)
        If CStr(vFileNameValue) <> ".gitignore" Then
            Call VBA.Kill(vDirectoryPath & "\" & CStr(vFileNameValue))
        End If
    Next

    ' Iterate through each component in the project.
    For Each vComponent In ThisWorkbook.VBProject.VBComponents
        ' Determine the file suffix to use.
        Select Case vComponent.Type
            Case VBIDE.vbext_ct_ClassModule, VBIDE.vbext_ct_Document
                vSuffix = "cls"
            Case VBIDE.vbext_ct_MSForm
                vSuffix = "frm"
            Case VBIDE.vbext_ct_StdModule
                vSuffix = "bas"
            Case Else
                vSuffix = VBA.vbNullString
        End Select

        ' Verify that an exportable component was encountered.
        If vSuffix <> VBA.vbNullString Then
            On Error Resume Next

            ' Attempt to export the file.
            vComponentPath = vDirectoryPath & "\" & vComponent.Name & "." & vSuffix
            Call vComponent.Export(vComponentPath)

            ' Report failure if it had occured.
            If VBA.Err.Number <> 0 Then
                Call VBA.MsgBox("Failed to export '" & vComponentPath & "'.")
                Call VBA.Err.Clear
            End If

            On Error GoTo 0
        End If
    Next
End Sub

Private Sub pActionOpenConfig()
    Call Application.Workbooks.Open(ConfigFilePath())
End Sub

Private Sub pActionUnprotectProject()
    ' Verify that the project
    If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
        On Error Resume Next

        ' Execute the unstable sequence of key presses.
        With Application
            Call .SendKeys("%{F11}", True)
            Call .Wait(500)
            Call .SendKeys("%(VP)", True)
            Call .Wait(500)
            Call .SendKeys("{DOWN}", True)
            Call .Wait(500)
            Call .SendKeys(MRuntimeParameters.vProjectPassword, True)
            Call .Wait(500)
            Call .SendKeys("{ENTER}", True)
            Call .SendKeys("{NUMLOCK}", True)
        End With

        On Error GoTo 0
    End If
End Sub

Private Sub pDisableSpeedMode()
    ' Disable event handling, formula calculation, screen updating and alert handling.
    With Application
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub

Private Sub pEnableSpeedMode()
    ' Enable alert handling, screen updating, formula calculation and event handling.
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
End Sub

Private Function pErrorMessage() As String
    pErrorMessage = vStoredErrorMessage
    If pErrorMessage = VBA.vbNullString Then
        pErrorMessage = MRuntimeParameters.vDefaultErrorMessage
    End If
End Function

Private Function pSerializedError( _
    ByRef vErrorMessage As String, _
    Optional ByRef vFullActionName As String = VBA.vbNullString _
) As String
    pSerializedError = "[DateTime] " & VBA.Format(VBA.Now, "yyyy/mm/dd Hh:Nn:Ss") & VBA.vbCrLf _
        & "[Computer and User] " & ComputerName() & " : " & Username() & VBA.vbCrLf _
        & "[Location] "

    If vFullActionName <> VBA.vbNullString Then
        pSerializedError = pSerializedError & vFullActionName & " : "
    End If

    pSerializedError = pSerializedError _
        & VBA.Err.Source & " : " & CStr(VBA.Err.Number) & VBA.vbCrLf _
        & "[Description] " & VBA.Err.Description & VBA.vbCrLf _
        & "[Message] " & vErrorMessage & VBA.vbCrLf
End Function

Private Sub pHandleError( _
    ByRef vFullActionName As String _
)
    ' Declare local variables.
    Dim vSerializedError As String
    Dim vErrorMessage As String
    Dim vRecipientDictionary As Dictionary

    ' Determine the error message.
    vErrorMessage = pErrorMessage()

    ' Serialize the error.
    vSerializedError = pSerializedError(vErrorMessage, vFullActionName)

    ' Determine whether the error has not been stored.
    If Not vIsErrorStored Then
        Debug.Print vSerializedError
        Debug.Assert Not pIsDebugMode()
    End If

    ' Report the error, while ignoring any errors in the reporter.
    If Not pIsDebugMode() Then
        On Error Resume Next
    End If

    ' Report the error to the user.
    If MRuntimeParameters.IsErrorMessageEnabled() Then
        Call VBA.MsgBox(vErrorMessage, vbCritical)
    End If

    ' Report the error to the administrator.
    If Not pIsDebugMode() Then
        ' Write the file to an error log.
        Call MUtilityFile.AppendData(ErrorFilePath(), vSerializedError & VBA.vbCrLf)

        ' Send an error message.
        Set vRecipientDictionary = New Dictionary
        vRecipientDictionary("To") = vReportErrorEmailAddress
        Call MUtilityMail.Send(vRecipientDictionary, _
            "[" & MRuntimeParameters.vProjectName & "] An unexpected error had occurred", _
            vSerializedError, vIsHtmlBody:=False)
    End If

    If Not pIsDebugMode() Then
        On Error GoTo 0
    End If

    ' Clear the stored error.
    vIsErrorStored = False
    vStoredErrorNumber = 0
    vStoredErrorSource = VBA.vbNullString
    vStoredErrorDescription = VBA.vbNullString
    vStoredErrorMessage = VBA.vbNullString
End Sub

Private Function pIsDebugMode() As Boolean
    If vDebugMode = EDebugMode.vUnknown Then
        If VBA.Environ$("APP_DEBUG_PASSWORD") = MRuntimeParameters.vProjectPassword Then
            vDebugMode = EDebugMode.vEnabled
        ElseIf VBA.Environ$("APP_DEBUG_PASSWORD") = VBA.vbNullString Then
            vDebugMode = EDebugMode.vDisabled
        Else
            vDebugMode = EDebugMode.vIncorrectPassword

            Call VBA.MsgBox("The debug password is incorrect.", vbCritical)

            If Application.Workbooks.Count > 1 Then
                Call ThisWorkbook.Close(False)
            Else
                Call Application.Quit
            End If
        End If
    End If
    pIsDebugMode = vDebugMode = EDebugMode.vEnabled
End Function

Private Function pIsInitializeOnStartupEnabled() As Boolean
    If pIsDebugMode() Then
        pIsInitializeOnStartupEnabled = _
            VBA.MsgBox("Do you wish to continue with the application's initialization?", vbYesNo) = vbYes
    Else
        pIsInitializeOnStartupEnabled = vDebugMode = EDebugMode.vDisabled
    End If
End Function

Private Sub pInitializeState()
    ' Declare local variables.
    Dim vConfigWorkbook As Workbook
    Dim vConfigTable As New CUtilityTable

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Collect environment parameters.
    vUsername = VBA.Environ$("USERNAME")
    vComputerName = VBA.Environ$("COMPUTERNAME")

    ' Open the config workbook.
    Set vConfigWorkbook = Application.Workbooks.Open(ConfigFilePath(), ReadOnly:=True, _
        Password:=MRuntimeParameters.vProjectPassword)

    ' Retrieve the mandatory runtime parameters.
    Call vConfigTable.UseHeaderRange(vConfigWorkbook.Worksheets("Main").Range("MainHeader"))
    With vConfigTable.GetValueToValueMap("Name", "Value")
        vReportErrorEmailAddress = .Item("ReportErrorEmailAddress")
        vDeployLocationPath = .Item("DeployLocationPath")
    End With

    Call MRuntimeParameters.InitializeState(vConfigWorkbook)

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Close the config workbook if opened.
    If Not (vConfigWorkbook Is Nothing) Then
        Call vConfigWorkbook.Close(False)
    End If

    ' Re-raise the error if needed.
    Call MRuntime.ReRaiseError

    ' Exit the procedure.
    Exit Sub

HandleError:
    ' Store the error for further handling.
    Call MRuntime.StoreError

    ' Resume to procedure termination.
    Resume Terminate:
End Sub

Private Sub pTerminateState()
    Call MRuntimeParameters.TerminateState
End Sub

Private Function pIsLessThanMaxRangeCount( _
    ByRef vRange As Range, _
    ByRef vMaxCount As Long _
) As Boolean
    On Error Resume Next
    pIsLessThanMaxRangeCount = vMaxCount > vRange.Count
    On Error GoTo 0
End Function

Public Sub SetErrorMessage( _
    ByVal vMessage As String _
)
    vStoredErrorMessage = vMessage
End Sub

Public Sub ClearErrorMessage()
    vStoredErrorMessage = VBA.vbNullString
End Sub

Public Sub RaiseError( _
    ByRef vSource As String, _
    ByRef vDescription As String, _
    Optional ByVal vMessage As String = VBA.vbNullString _
)
    ' Store the error message.
    If vMessage <> VBA.vbNullString Then
        vStoredErrorMessage = vMessage
    End If

    ' Raise the error with the correct number and description.
    Call VBA.Err.Raise(vApplicationDefinedErrorNumber, vSource, vDescription)
End Sub

Public Sub ClearError()
    ' Reset the stored error flag.
    vIsErrorStored = False
    Call VBA.Err.Clear
End Sub

Public Sub StoreError()
    ' Reset the stored error flag.
    vIsErrorStored = True

    ' Store the current error parameters.
    vStoredErrorNumber = VBA.Err.Number
    vStoredErrorSource = VBA.Err.Source
    vStoredErrorDescription = VBA.Err.Description

    ' Output the error message to the debug window.
    Debug.Print pSerializedError(pErrorMessage())

    ' Start debugging if in debug mode.
    Debug.Assert Not pIsDebugMode()
End Sub

Public Sub ReRaiseError()
    ' Verify that an error is stored.
    If vIsErrorStored Then
        ' Reset the stored error flag.
        vIsErrorStored = False

        ' ReRaise an error with the stored error parameters.
        Call VBA.Err.Raise(vStoredErrorNumber, vStoredErrorSource, vStoredErrorDescription)
    End If
End Sub

Public Sub HideInterface()
    ' Hide the status bar and the formula bar.
    With Application
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .DisplayFullScreen = True
    End With
End Sub

Public Sub ShowInterface()
    ' Show the status bar and the formula bar.
    With Application
        .DisplayFullScreen = False
        .DisplayStatusBar = True
        .DisplayFormulaBar = True
    End With
End Sub

Public Sub Refresh()
    ' Recalculate and prompt event execution.
    Call Application.Calculate
    Call VBA.DoEvents
End Sub

Public Sub Quit()
    ' Ensures that the application is closed when the current action execution ends.
    vIsQuitRequested = True
End Sub

Public Function ReportErrorEmailAddress() As String
    ReportErrorEmailAddress = vReportErrorEmailAddress
End Function

Public Function DeployLocationPath() As String
    DeployLocationPath = vDeployLocationPath
End Function

Public Function Username() As String
    Username = vUsername
End Function

Public Function ComputerName() As String
    ComputerName = vComputerName
End Function

Public Function ConfigFilePath() As String
    ConfigFilePath = ThisWorkbook.Path & "\Config.xlsx"
End Function

Public Function ErrorFilePath() As String
    ErrorFilePath = ThisWorkbook.Path & "\Error.log"
End Function

Public Sub Execute( _
    ByRef vObjectName As String, _
    ByRef vActionName As String, _
    Optional ByVal vIsActiveLockRequired As Boolean = True, _
    Optional ByVal vIsDebugRequired As Boolean = False, _
    Optional ByVal vIsSpeedModeEnabled As Boolean = True, _
    Optional ByVal vEventHandling As EEventHandling = vNone, _
    Optional ByRef vEventParameter As Variant = Null, _
    Optional ByVal vIsIgnored As Boolean = False _
)
    ' Declare local variables.
    Dim vParameterRange As Range
    Dim vFullActionName As String
    Dim vHasErrorOccurred As Boolean
    Dim vIsSpeedModeRequired As Boolean
    Dim vIsAuthorized As Boolean
    Dim vIsTerminatingState As Boolean
    Dim vIsQuiting As Boolean

    ' Execute special event handling.
    vFullActionName = vObjectName & "." & vActionName
    Select Case vEventHandling
        Case EEventHandling.vWorkbookOpen
            ' Prompt for disabling the initialization in debug mode.
            If Not pIsInitializeOnStartupEnabled() Then
                Exit Sub
            End If
        Case EEventHandling.vWorkbookBeforeClose
            ' Disable the save prompt when not in debug mode.
            If pIsDebugMode() Then
                If ThisWorkbook.VBProject.Protection = vbext_pp_none Then
                    Call pActionExportProject
                End If
            Else
                ThisWorkbook.Saved = True
            End If
        Case EEventHandling.vWorksheetChange
            Set vParameterRange = vEventParameter
            If Not pIsLessThanMaxRangeCount(vParameterRange, _
                MRuntimeParameters.MaxChangeRangeCount(vObjectName, vActionName)) _
            Then
                Exit Sub
            End If
        Case EEventHandling.vWorksheetSelectionChange
            Set vParameterRange = vEventParameter
            If Not pIsLessThanMaxRangeCount(vParameterRange, _
                MRuntimeParameters.MaxSelectionChangeRangeCount(vObjectName, vActionName)) _
            Then
                Exit Sub
            End If
    End Select

    ' Check if the action is to be ignored.
    If vIsIgnored Then
        Exit Sub
    End If

    ' Check the action execution lock.
    If vIsActiveLockRequired Then
        If vIsActive Then
            Exit Sub
        End If
        vIsActive = True
    End If

    ' Configure error handling.
    On Error GoTo HandleError:

    ' Reset the quit flag.
    vIsQuitRequested = False

    ' Set the loading mouse pointer.
    If vIsSpeedModeEnabled And (Not pIsDebugMode()) Then
        Application.Cursor = xlWait
    End If

    ' Enable speed mode.
    If vIsSpeedModeEnabled Then
        Call pEnableSpeedMode
    End If

    ' Initialize the state if not initialized.
    If Not vIsInitialized Then
        Call pInitializeState
        vIsInitialized = True
    End If

    ' Verify that the file is run from the deployment location if not in debug mode.
    If vEventHandling = EEventHandling.vWorkbookOpen Then
        If (Not pIsDebugMode()) And (vDeployLocationPath <> VBA.vbNullString) Then
            If VBA.LCase(vDeployLocationPath) <> VBA.LCase(ThisWorkbook.Path) Then
                Call RaiseError("MRuntime.Execute", "The application has been run from '" _
                    & ThisWorkbook.Path & "', but can only be run from '" & vDeployLocationPath & "'.")
            End If
        End If
    End If

    ' Attempt to execute the selected action.
    vIsAuthorized = True
    If vIsDebugRequired Then
        vIsAuthorized = pIsDebugMode()
    End If
    If vIsAuthorized Then
        Select Case vFullActionName
            Case "ThisWorkbook.Deploy"
                Call pActionDeploy
            Case "ThisWorkbook.ExportProject"
                Call pActionExportProject
            Case "ThisWorkbook.OpenConfig"
                Call pActionOpenConfig
            Case "ThisWorkbook.UnprotectProject"
                Call pActionUnprotectProject
            Case Else
                If Not MRuntimeParameters.Execute(vFullActionName, vEventParameter) Then
                    Call RaiseError("MRuntime.Execute", "The action '" & vFullActionName & "' does not have a registered handler.")
                End If
        End Select
    Else
        Call VBA.MsgBox("Can only be run in debug mode.", vbExclamation)
    End If

TerminateState:
    ' Set terminating state flag.
    vIsTerminatingState = True

    ' Terminate the state if the application is quitting or an error had occurred.
    If vIsQuitRequested Or vHasErrorOccurred Then
        vIsInitialized = False
        Call pTerminateState
    End If

Terminate:
    ' Reset error handling
    On Error GoTo 0

    ' Disable speed mode and refresh.
    If vIsSpeedModeEnabled Then
        Call pDisableSpeedMode
        Call Refresh
    End If

    ' Reset the loading mouse pointer.
    If vIsSpeedModeEnabled And (Not pIsDebugMode()) Then
        Application.Cursor = xlDefault
    End If

    ' Quit the application if the application is quitting or not in debug mode when an error occurred.
    vIsQuiting = vIsQuitRequested
    If Not vIsQuiting Then
        vIsQuiting = Not pIsDebugMode() And vHasErrorOccurred
    End If
    If vIsQuiting Then
        If Application.Workbooks.Count > 1 Then
            Call ThisWorkbook.Close(False)
        Else
            Call Application.Quit
        End If
    End If

    ' Release the action execution lock.
    If vIsActiveLockRequired Then
        vIsActive = False
    End If

    ' Exit the procedure.
    Exit Sub

HandleError:
    ' Set the error flag.
    vHasErrorOccurred = True

    ' Handle error reporting.
    Call pHandleError(vFullActionName)

    ' Terminate error handling.
    If vIsTerminatingState Then
        Resume Terminate:
    End If
    Resume TerminateState:
End Sub
