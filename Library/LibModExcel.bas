Attribute VB_Name = "LibModExcel"
Option Explicit
Option Private Module

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

Public Sub Refresh()
    ' Recalculate and prompt event execution.
    Call Application.Calculate
    Call VBA.DoEvents
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

Attribute VB_Name = "LibModWorkbook"
Option Explicit
Option Private Module

' Requires MRuntime

Private Const vSubscriptOutOfRangeErrorNumber As Long = 9

' TODO: Refactor to not use error.
Public Function SheetExists( _
    ByRef vWorkbook As Workbook, _
    ByRef vSheetName As String _
) As Boolean
    ' Declare local variables.
    Dim vWorksheet As Worksheet

    ' Setup error handling.
    On Error GoTo HandleError:

    ' Determine the sheet's existance.
    Set vWorksheet = vWorkbook.Worksheets(vSheetName)
    SheetExists = True

Terminate:
    ' Reset error handling.
    On Error GoTo 0

    ' Re-raise the error if needed.
    Call MRuntime.ReRaiseError

    ' Exit the procedure.
    Exit Function

HandleError:
    ' Store the error for further handling.
    If VBA.Err.Number <> vSubscriptOutOfRangeErrorNumber Then
        Call MRuntime.StoreError
    End If

    ' Resume to procedure termination.
    Resume Terminate:
End Function

Attribute VB_Name = "LibModRange"
Option Explicit
Option Private Module

' Requires MRuntime

Public Function IsEmpty( _
    ByRef vRange As Range _
) As Boolean
    ' Declare local variables.
    Dim vCellRange As Range

    ' Set the initial result.
    IsEmpty = True

    ' Search for a non empty cell value within the range.
    For Each vCellRange In vRange
        If IsEmpty Then
            If CStr(vCellRange.Value2) <> VBA.vbNullString Then
                IsEmpty = False
                Exit Function
            End If
        End If
    Next
End Function

Public Function GetLastColumnRange( _
    ByRef vColumnRange As Range _
) As Range
    ' Verify that the given range spans a single row.
    If vColumnRange.Columns.Count <> 1 Then
        Call MRuntime.RaiseError("MUtilityRange.GetLastColumnRange", "The submitted range '" _
            & vColumnRange.Address() & "' must have exactly one column.")
    End If

    ' Set the initial result.
    Set GetLastColumnRange = vColumnRange

    ' Verify if the current row range is empty.
    If IsEmpty(GetLastColumnRange) Then
        Exit Function
    End If

    ' Verify if the next row range is empty.
    If IsEmpty(GetLastColumnRange.Offset(0, 1)) Then
        Exit Function
    End If

    ' Skip to the bottom.
    Set GetLastColumnRange = GetLastColumnRange.Offset(0, _
        GetLastColumnRange.End(xlToRight).Column - GetLastColumnRange.Column)

    ' Iterate downward until an empty row is encountered.
    Do While Not IsEmpty(GetLastColumnRange.Offset(0, 1))
        Set GetLastColumnRange = GetLastColumnRange.Offset(0, 1)
    Loop
End Function

Public Function GetLastRowRange( _
    ByRef vRowRange As Range _
) As Range
    ' Verify that the given range spans a single row.
    If vRowRange.Rows.Count <> 1 Then
        Call MRuntime.RaiseError("MUtilityRange.GetLastRowRange", "The submitted range '" _
            & vRowRange.Address() & "' must have exactly one row.")
    End If

    ' Set the initial result.
    Set GetLastRowRange = vRowRange

    ' Verify if the current row range is empty.
    If IsEmpty(GetLastRowRange) Then
        Exit Function
    End If

    ' Verify if the next row range is empty.
    If IsEmpty(GetLastRowRange.Offset(1)) Then
        Exit Function
    End If

    ' Skip to the bottom.
    Set GetLastRowRange = GetLastRowRange.Offset(GetLastRowRange.End(xlDown).Row - GetLastRowRange.Row)

    ' Iterate downward until an empty row is encountered.
    Do While Not IsEmpty(GetLastRowRange.Offset(1))
        Set GetLastRowRange = GetLastRowRange.Offset(1)
    Loop
End Function

