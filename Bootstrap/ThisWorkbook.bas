Option Explicit

' Requires Runtime

Private Const vDefaultWidth As Long = 800
Private Const vDefaultHeight As Long = 400

Private Sub pInitialize()
    ' Check whether the application is running in background mode.
    If Runtime.IsBackgroundModeEnabled() Then
        ' Execute the default navigate path.
        Call Runtime.Navigate(Runtime.NavigatePath())
    Else
        ' Show the main user form.
        Call ThisUserForm.Show
    End If
End Sub

Private Sub Workbook_BeforeClose( _
    Cancel As Boolean _
)
    If _
        (Not Runtime.IsDebugModeEnabled()) _
        Or Runtime.IsDeployDebugModeEnabled() _
    Then
        Me.Saved = True
    End If
End Sub

Private Sub Workbook_Open()
    ' Load the current excel application instance.
    With Application
        ' Check whether the application is visible.
        If .Visible Then
            ' Reset the dimensions.
            .Width = vDefaultWidth
            .Height = vDefaultHeight

            ' Set the title and icon of the window.
            .ActiveWindow.Caption = vbNullString
            .Caption = Runtime.ProjectName()
            Call Runtime.SetActiveWindowIcon
        End If

        ' Disable unnecessary activities.
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    ' If the application is in debug mode, do not continue.
    If Runtime.IsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Initialize the application.
    Call pInitialize

    ' Close the excel application instance.
    Call Application.Quit
End Sub

Public Sub Initialize()
    ' If the application is not in debug mode, do not continue.
    If Not Runtime.IsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Initialize the application.
    Call pInitialize
End Sub

Public Sub InitializeWithPath()
    ' If the application is not in debug mode, do not continue.
    If Not Runtime.IsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Set the navigate path environment variable.
    Runtime.WScriptShell().Environment("PROCESS")("APP_NAVIGATE_PATH") = InputBox("Enter the path to navigate to")

    ' Initialize the application.
    Call pInitialize
End Sub
