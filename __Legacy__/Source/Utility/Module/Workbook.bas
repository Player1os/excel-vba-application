Option Explicit
Option Private Module

' Requires MRuntime

Private Const vSubscriptOutOfRangeErrorNumber As Long = 9

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
