Option Explicit

Private Sub pInitialize()
    ' Load the current excel application instance.
    With Application
        ' Set the title and icon of the window.
        .ActiveWindow.Caption = vbNullString
        .Caption = Runtime.ProjectName()
        Call Runtime.SetActiveWindowIcon

        ' Configure to run in speed mode.
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    ' Show the main user form.
    Call ThisUserForm.Show
End Sub

Private Sub Workbook_Open()
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
