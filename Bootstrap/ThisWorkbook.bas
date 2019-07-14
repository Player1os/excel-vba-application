Option Explicit

Private Function pIsDebugModeEnabled()
    pIsDebugModeEnabled = VBA.Environ$("APP_IS_DEBUG_MODE_ENABLED") = "TRUE"
End Function

Private Sub pInitialize()
    ' Configure the excel application instance to run in speed mode.
    With Application
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
    If pIsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Initialize the application.
    Call pInitialize

    ' Close the excel application instance.
    Call Application.Quit
End Sub

Public Sub Initialize()
    ' If the application is not in debug mode, do not continue.
    If Not pIsDebugModeEnabled() Then
        Exit Sub
    End If

    ' Initialize the application.
    Call pInitialize
End Sub